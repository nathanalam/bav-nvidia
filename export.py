"""Create a formatted Excel workbook from the curated filings SQLite database."""

from __future__ import annotations

import argparse
import sqlite3
from datetime import datetime
from pathlib import Path

import pandas as pd
from statement_layout import TOTAL_LINE_ITEMS, order_statement_frame

DB_PATH = Path(__file__).parent / "filings.sqlite3"
DEFAULT_OUTPUT = Path(__file__).parent / "10k_financial_statements.xlsx"
STATEMENTS = ["Balance sheet", "Income statement", "Cash flow statement"]


def load(sql: str, params=()) -> pd.DataFrame:
    with sqlite3.connect(DB_PATH) as conn:
        return pd.read_sql_query(sql, conn, params=params)


def style_sheet(writer: pd.ExcelWriter, name: str, frame: pd.DataFrame, widths: dict[str, int] | None = None) -> None:
    workbook = writer.book
    worksheet = writer.sheets[name]
    header = workbook.add_format({"bold": True, "font_color": "white", "bg_color": "#17365D", "border": 0})
    title = workbook.add_format({"bold": True, "font_size": 16, "font_color": "#17365D"})
    total = workbook.add_format({"bold": True, "top": 1})
    worksheet.write(0, 0, name, title)
    worksheet.set_row(0, 24)
    worksheet.set_row(1, 8)
    for column_index, column in enumerate(frame.columns):
        worksheet.write(2, column_index, column, header)
    worksheet.freeze_panes(3, 1)
    worksheet.autofilter(2, 0, max(len(frame) + 2, 2), max(len(frame.columns) - 1, 0))
    for index, column in enumerate(frame.columns):
        width = (widths or {}).get(column, min(max(len(str(column)) + 2, int(frame[column].astype(str).str.len().quantile(.9)) + 2 if len(frame) else 12), 42))
        worksheet.set_column(index, index, width)
    for row_index, label in enumerate(frame["Line item"], start=3):
        if label in TOTAL_LINE_ITEMS:
            worksheet.set_row(row_index, None, total)


def statement_frame(company: str, category: str) -> pd.DataFrame:
    frame = load("""SELECT m.metric AS 'Line item', f.fiscal_year AS filing_year, m.period, m.value, m.source_page AS 'Source page'
        FROM metrics m JOIN filings f ON f.id=m.filing_id
        WHERE f.company=? AND m.category=?
        ORDER BY f.fiscal_year DESC, m.source_page, m.id""", (company, category))
    if frame.empty:
        return frame
    frame["period_year"] = frame["period"].str.extract(r"(\d{4})")[0].astype("Int64")
    frame["period_priority"] = (frame["filing_year"] == frame["period_year"]).astype(int)
    frame = frame.sort_values(["Line item", "period", "period_priority", "filing_year"], ascending=[True, True, False, False])
    frame = frame.drop_duplicates(["Line item", "period"], keep="first")
    years = sorted((int(year) for year in frame["period_year"].dropna().unique()), reverse=True)
    frame["period"] = frame["period_year"].map(lambda year: f"FY {int(year)}")
    wide = frame.pivot(index="Line item", columns="period", values="value").reset_index()
    periods = [f"FY {year}" for year in years]
    return order_statement_frame(wide[["Line item"] + periods], category)


def create_workbook(output: Path) -> Path:
    filings = load("SELECT id,company,filename,filing_date,fiscal_year,pages,parsed_at,sha256,path FROM filings ORDER BY company,fiscal_year")

    output.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output, engine="xlsxwriter", datetime_format="yyyy-mm-dd") as writer:
        workbook = writer.book
        workbook.set_properties({"title": "10-K Financial Statements", "subject": "Curated financial statement data", "author": "10-K Financial Explorer", "comments": f"Exported {datetime.now().isoformat(timespec='seconds')}"})
        number_format = workbook.add_format({"num_format": "#,##0.00;[Red](#,##0.00)"})
        for company in sorted(filings.company.unique()):
            for category in STATEMENTS:
                frame = statement_frame(company, category)
                if frame.empty:
                    continue
                name = f"{company} - {category}"[:31]
                frame.to_excel(writer, sheet_name=name, index=False, startrow=2)
                style_sheet(writer, name, frame, {"Line item": 42})
                worksheet = writer.sheets[name]
                if len(frame.columns) > 1:
                    worksheet.set_column(1, len(frame.columns) - 1, 18, number_format)
    return output


def main() -> None:
    parser = argparse.ArgumentParser(description="Export curated 10-K data to one formatted Excel workbook.")
    parser.add_argument("-o", "--output", type=Path, default=DEFAULT_OUTPUT, help=f"Output path (default: {DEFAULT_OUTPUT.name})")
    args = parser.parse_args()
    if not DB_PATH.exists():
        raise SystemExit(f"Database not found: {DB_PATH}")
    result = create_workbook(args.output)
    print(f"Created {result} ({result.stat().st_size:,} bytes)")


if __name__ == "__main__":
    main()
