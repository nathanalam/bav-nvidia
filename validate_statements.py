"""Validate that the curated SQLite database has all core statements per filing."""

from __future__ import annotations

import sqlite3
import sys
from pathlib import Path


DB_PATH = Path(__file__).parent / "filings.sqlite3"
REQUIRED = ("Balance sheet", "Income statement", "Cash flow statement")


def main() -> int:
    if not DB_PATH.exists():
        print(f"FAIL: database not found: {DB_PATH}")
        return 1
    conn = sqlite3.connect(DB_PATH)
    filings = conn.execute("""
        SELECT id, company, fiscal_year
        FROM filings
        ORDER BY company, fiscal_year
    """).fetchall()
    missing = []
    print("company | fiscal year | balance sheet | income statement | cash flow statement")
    print("---|---:|---:|---:|---:")
    for filing_id, company, year in filings:
        counts = dict(conn.execute("""
            SELECT category, COUNT(*)
            FROM metrics
            WHERE filing_id = ?
            GROUP BY category
        """, (filing_id,)).fetchall())
        present = [counts.get(category, 0) for category in REQUIRED]
        print(f"{company} | {year} | {present[0]} | {present[1]} | {present[2]}")
        for category, count in zip(REQUIRED, present):
            if count == 0:
                missing.append((company, year, category))
    conn.close()
    if missing:
        print("\nMISSING STATEMENTS")
        for company, year, category in missing:
            print(f"- {company} FY {year}: {category}")
        return 1
    print(f"\nPASS: all {len(filings)} filings have all {len(REQUIRED)} core statements.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
