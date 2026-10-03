"""Create a formatted Excel workbook from the curated filings SQLite database."""

from __future__ import annotations

import argparse
import sqlite3
from datetime import datetime
from pathlib import Path

import pandas as pd

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


def statement_frame(company: str, year: int, category: str) -> pd.DataFrame:
    frame = load("""SELECT m.metric AS 'Line item', m.period, m.value, m.source_page AS 'Source page'
        FROM metrics m JOIN filings f ON f.id=m.filing_id
        WHERE f.company=? AND f.fiscal_year=? AND m.category=?
        ORDER BY m.source_page, m.id""", (company, int(year), category))
    if frame.empty:
        return frame
    frame = frame.drop_duplicates(["Line item", "period"], keep="first")
    wide = frame.pivot(index="Line item", columns="period", values="value").reset_index()
    periods = list(dict.fromkeys(frame["period"].tolist()))
    pages = frame.groupby("Line item", as_index=False)["Source page"].min()
    return wide.merge(pages, on="Line item", how="left")[["Line item"] + periods + ["Source page"]]


def create_workbook(output: Path) -> Path:
    filings = load("SELECT id,company,filename,filing_date,fiscal_year,pages,parsed_at,sha256,path FROM filings ORDER BY company,fiscal_year")
    metrics = load("""SELECT m.id,m.filing_id,f.company,f.fiscal_year,m.category,m.metric,m.period,m.value,m.unit,m.source_page,m.source_text
        FROM metrics m JOIN filings f ON f.id=m.filing_id ORDER BY f.company,f.fiscal_year,m.category,m.source_page,m.id""")
    pages = load("""SELECT p.id,p.filing_id,f.company,f.fiscal_year,p.page_number,p.text
        FROM pages p JOIN filings f ON f.id=p.filing_id ORDER BY f.company,f.fiscal_year,p.page_number""")
    coverage = metrics.groupby(["company", "fiscal_year", "category"], as_index=False).agg(Rows=("id", "count"))
    coverage_pivot = coverage.pivot_table(index=["company", "fiscal_year"], columns="category", values="Rows", fill_value=0).reset_index()
    coverage_pivot.columns.name = None

    output.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output, engine="xlsxwriter", datetime_format="yyyy-mm-dd") as writer:
        workbook = writer.book
        workbook.set_properties({"title": "10-K Financial Statements", "subject": "Curated financial statement data", "author": "10-K Financial Explorer", "comments": f"Exported {datetime.now().isoformat(timespec='seconds')}"})
        overview = pd.DataFrame({"Field": ["Exported at", "Database", "Filings", "Pages", "Metrics", "Companies"], "Value": [datetime.now().isoformat(timespec="seconds"), str(DB_PATH), len(filings), len(pages), len(metrics), filings["company"].nunique()]})
        overview.to_excel(writer, sheet_name="Read me", index=False, startrow=2)
        style_sheet(writer, "Read me", overview, {"Field": 20, "Value": 58})
        coverage_pivot.to_excel(writer, sheet_name="Coverage", index=False, startrow=2)
        style_sheet(writer, "Coverage", coverage_pivot, {"company": 16, "fiscal_year": 14})

        for company in sorted(filings.company.unique()):
            for year in sorted(filings.loc[filings.company == company, "fiscal_year"].dropna().astype(int).unique()):
                for category in STATEMENTS:
                    frame = statement_frame(company, year, category)
                    if frame.empty:
                        continue
                    name = f"{company[:8]}_{year}_{category[:5]}"[:31]
                    frame.to_excel(writer, sheet_name=name, index=False, startrow=2)
                    style_sheet(writer, name, frame, {"Line item": 42, "Source page": 13})
                    worksheet = writer.sheets[name]
                    number_format = workbook.add_format({"num_format": "#,##0.00;[Red](#,##0.00)"})
                    if len(frame.columns) > 2:
                        worksheet.set_column(1, len(frame.columns) - 2, 16, number_format)

        metrics.to_excel(writer, sheet_name="All metrics", index=False, startrow=2)
        style_sheet(writer, "All metrics", metrics, {"source_text": 70, "metric": 32, "category": 25})
        filings.to_excel(writer, sheet_name="Filings", index=False, startrow=2)
        style_sheet(writer, "Filings", filings, {"path": 55, "sha256": 68, "filename": 42})
        pages.to_excel(writer, sheet_name="Source pages", index=False, startrow=2)
        style_sheet(writer, "Source pages", pages, {"text": 100, "company": 16})
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
