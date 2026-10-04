from __future__ import annotations

import io
import json
import re
import sqlite3
from datetime import datetime, timezone
from pathlib import Path

import pandas as pd
import streamlit as st
from statement_layout import UNIT_CAPTION, order_statement_frame
from reportlab.lib import colors
from reportlab.lib.pagesizes import landscape, letter
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

DB_PATH = Path(__file__).parent / "filings.sqlite3"
STATEMENT_CATEGORIES = ["Balance sheet", "Income statement", "Cash flow statement"]


def connect() -> sqlite3.Connection:
    return sqlite3.connect(DB_PATH)


def init_db(conn: sqlite3.Connection) -> None:
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS filings (
            id INTEGER PRIMARY KEY, company TEXT NOT NULL, filename TEXT NOT NULL,
            path TEXT NOT NULL UNIQUE, filing_date TEXT, fiscal_year INTEGER,
            pages INTEGER, sha256 TEXT NOT NULL, parsed_at TEXT NOT NULL, raw_text TEXT
        );
        CREATE TABLE IF NOT EXISTS pages (
            id INTEGER PRIMARY KEY, filing_id INTEGER NOT NULL, page_number INTEGER,
            text TEXT, FOREIGN KEY(filing_id) REFERENCES filings(id) ON DELETE CASCADE
        );
        CREATE TABLE IF NOT EXISTS metrics (
            id INTEGER PRIMARY KEY, filing_id INTEGER NOT NULL, category TEXT,
            metric TEXT NOT NULL, period TEXT, value REAL, unit TEXT, source_page INTEGER,
            source_text TEXT, justification TEXT NOT NULL DEFAULT '',
            FOREIGN KEY(filing_id) REFERENCES filings(id) ON DELETE CASCADE
        );
    """)
    metric_columns = {row[1] for row in conn.execute("PRAGMA table_info(metrics)")}
    if "justification" not in metric_columns:
        conn.execute("ALTER TABLE metrics ADD COLUMN justification TEXT NOT NULL DEFAULT ''")
    conn.commit()


def query_df(sql: str, params=()) -> pd.DataFrame:
    conn = connect()
    try:
        return pd.read_sql_query(sql, conn, params=params)
    finally:
        conn.close()


def database_json_export() -> bytes:
    """Export every database table, including page-level source text, as JSON."""
    conn = connect()
    conn.row_factory = sqlite3.Row
    payload = {
        "format": "10k-financial-statements",
        "version": 1,
        "exported_at": datetime.now(timezone.utc).isoformat(),
        "tables": {},
    }
    for table in ("filings", "pages", "metrics"):
        rows = conn.execute(f"SELECT * FROM {table} ORDER BY id").fetchall()
        payload["tables"][table] = [dict(row) for row in rows]
    conn.close()
    return json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8")


def reconstructed_statement(company: str, year: int, category: str) -> pd.DataFrame:
    frame = query_df(
        """SELECT m.metric, m.period, m.value, m.source_page
           FROM metrics m JOIN filings f ON f.id = m.filing_id
           WHERE f.company = ? AND f.fiscal_year = ? AND m.category = ?
           ORDER BY m.source_page, m.id""",
        (company, int(year), category),
    )
    if frame.empty:
        return pd.DataFrame(columns=["Line item", "Source page"])
    frame = frame.drop_duplicates(["metric", "period"], keep="first")
    table = frame.pivot(index="metric", columns="period", values="value").reset_index().rename(columns={"metric": "Line item"})
    periods = list(dict.fromkeys(frame["period"].tolist()))
    pages = frame.groupby("metric", as_index=False)["source_page"].min().rename(columns={"metric": "Line item", "source_page": "Source page"})
    result = table.merge(pages, on="Line item", how="left")[["Line item"] + periods + ["Source page"]]
    return order_statement_frame(result, category)


def excel_export(company: str, year: int, statements: dict[str, pd.DataFrame]) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book
        title = workbook.add_format({"bold": True, "font_size": 16, "font_color": "#17365D"})
        number = workbook.add_format({"num_format": "#,##0.00;[Red](#,##0.00)"})
        for category, frame in statements.items():
            sheet = category.replace(" statement", "").replace(" ", "_")[:31]
            frame.to_excel(writer, sheet_name=sheet, index=False, startrow=2)
            ws = writer.sheets[sheet]
            ws.write(0, 0, f"{company} — {category} — FY {year}", title)
            ws.write(1, 0, UNIT_CAPTION)
            ws.freeze_panes(3, 1)
            ws.set_column(0, 0, 38)
            if len(frame.columns) > 2:
                ws.set_column(1, len(frame.columns) - 2, 16, number)
            ws.set_column(len(frame.columns) - 1, len(frame.columns) - 1, 12)
            ws.autofilter(2, 0, max(len(frame) + 2, 2), len(frame.columns) - 1)
    return output.getvalue()


def pdf_export(company: str, year: int, statements: dict[str, pd.DataFrame]) -> bytes:
    output = io.BytesIO()
    doc = SimpleDocTemplate(output, pagesize=landscape(letter), rightMargin=.35 * inch, leftMargin=.35 * inch, topMargin=.35 * inch, bottomMargin=.35 * inch)
    styles = getSampleStyleSheet()
    story = [Paragraph(f"{company} — Financial Statements — FY {year}", styles["Title"]), Paragraph(UNIT_CAPTION, styles["Italic"]), Spacer(1, .12 * inch)]
    for category, frame in statements.items():
        story.append(Paragraph(category, styles["Heading2"]))
        if frame.empty:
            story.append(Paragraph("No curated rows available.", styles["BodyText"]))
            continue
        display = frame.copy().fillna("")
        for column in display.columns[1:-1]:
            display[column] = display[column].map(lambda x: f"{x:,.2f}" if isinstance(x, (int, float)) else x)
        data = [list(display.columns)] + display.astype(str).values.tolist()
        table = Table(data, repeatRows=1, colWidths=[2.5 * inch] + [1.0 * inch] * (len(display.columns) - 2) + [.7 * inch])
        table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#17365D")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 7),
            ("GRID", (0, 0), (-1, -1), .25, colors.HexColor("#B7C9D6")),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F2F6FA")]),
            ("ALIGN", (1, 1), (-1, -1), "RIGHT"),
        ]))
        story.extend([table, Spacer(1, .16 * inch)])
    doc.build(story)
    return output.getvalue()


def app() -> None:
    st.set_page_config(page_title="Curated 10-K Financial Statements", page_icon="📊", layout="wide")
    st.title("10-K Financial Statements")
    st.caption("Curated statement rows from the attached 10-K filings; every row retains its source page.")
    conn = connect()
    init_db(conn)
    conn.close()
    filings = query_df("SELECT company, fiscal_year FROM filings ORDER BY company, fiscal_year DESC")
    if filings.empty:
        st.error("The curated database is missing. filings.sqlite3 must be present beside the app.")
        return
    with st.sidebar:
        st.subheader("Data export")
        st.download_button(
            "Download complete database as JSON",
            data=database_json_export(),
            file_name="10k_financial_statements.json",
            mime="application/json",
            use_container_width=True,
        )
    controls = st.columns(3)
    company = controls[0].selectbox("Company", sorted(filings.company.unique()))
    years = sorted((int(year) for year in filings.loc[filings.company == company, "fiscal_year"].dropna().unique()), reverse=True)
    year = controls[1].selectbox("Fiscal year", years)
    category = controls[2].selectbox("Statement", STATEMENT_CATEGORIES)
    frame = reconstructed_statement(company, year, category)
    st.subheader(f"{company} — {category} — FY {year}")
    st.caption(f"{UNIT_CAPTION}. Source page is the filing page containing the curated row.")
    if frame.empty:
        st.info("No curated rows are available for this company, year, and statement.")
    else:
        st.dataframe(frame.style.format({c: "{:,.2f}" for c in frame.columns[1:-1]}, na_rep="—"), use_container_width=True, hide_index=True, height=600)
    statements = {name: reconstructed_statement(company, year, name) for name in STATEMENT_CATEGORIES}
    safe = re.sub(r"[^A-Za-z0-9_-]+", "_", f"{company}_{year}")
    left, right = st.columns(2)
    left.download_button("Download Excel workbook", excel_export(company, year, statements), f"{safe}_financial_statements.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    right.download_button("Download PDF report", pdf_export(company, year, statements), f"{safe}_financial_statements.pdf", "application/pdf", use_container_width=True)


if __name__ == "__main__":
    app()
