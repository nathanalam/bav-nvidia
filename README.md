 # 10-K Financial Explorer

 A Streamlit app for parsing the PDFs under `10k filings/` into `filings.sqlite3` and exploring balance sheets, income statements, cash flows, market-value disclosures, operating ratios, forecasts, and source text.

 ## Run

 ```powershell
 uv sync
 uv run streamlit run main.py
 ```

 Open the local URL shown by Streamlit. The app is read-only: it opens the curated `filings.sqlite3` database and does not parse PDFs at runtime.

 ## SQLite tables

 - `filings`: one record per PDF and its full extracted text
 - `pages`: page-level source text for auditability
 - `metrics`: normalized line items, category, value, and source page

 The statement dataset is curated from the statement pages in the attached filings. Important figures retain their source page in the app and in the exported reports.
