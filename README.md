 # 10-K Financial Explorer

 A Streamlit app for parsing the PDFs under `10k filings/` into `filings.sqlite3` and exploring balance sheets, income statements, cash flows, market-value disclosures, operating ratios, forecasts, and source text.

 ## Run

 ```powershell
 uv sync
 uv run streamlit run main.py
 ```

 Open the local URL shown by Streamlit. The app is read-only: it opens the curated `filings.sqlite3` database and does not parse PDFs at runtime.

 Use **Download complete database as JSON** in the sidebar to export all `filings`, `pages`, and `metrics` rows, including source text and page citations, into `10k_financial_statements.json`.

 To create one formatted Excel workbook from the database:

 ```powershell
 uv run python export.py
 ```

 This writes `10k_financial_statements.xlsx` with one sheet per company/statement pair. Each sheet contains all available fiscal years as columns. Use `uv run python export.py --output my_export.xlsx` for a different path.

 ## Extraction evaluation methodology

 `eval.py` evaluates each extracted metric against the filing evidence retained in `filings.sqlite3`. It uses the page-level source text and source-page citation associated with each metric; no value is accepted based on the model's confidence alone.

 For every extracted value, the evaluator:

 1. Checks deterministically that a textual form of the extracted number appears in the cited evidence. Missing numeric evidence is an automatic failure because decision models are not reliable numeric calculators.
 2. Sends one Decisions API request to the pinned OpenRouter model. The request asks two independent yes/no (`noul`) questions: whether the value is supported for the same metric and period, and whether it is likely hallucinated or unsupported. An optional extractor justification is included as a claim for the evaluator to check against the filing, but it never replaces filing evidence.
 3. Reports the model's support probability as the value's confidence and reports the hallucination probability separately.
 4. Passes a value only when numeric evidence is present, support confidence is at least 90%, and the hallucination assessment is below 50%.

 The complete evaluation passes only when every value passes. Any API error, missing evidence, confidence below 90%, or likely hallucination makes the process exit with status 1.

 Values are evaluated concurrently with a bounded thread pool because each Decisions API request is independent. The default is eight workers; use `--workers 1` for sequential execution, or set `EVAL_WORKERS` to tune concurrency. Transient rate-limit and server errors are retried up to twice with backoff.

 The evaluator reads the OpenRouter key from `C:\Users\Nathan\.openrouter\api_key.txt` by default without printing it. It uses the pinned `typesafe/jev-1.13-20260917` model by default; override it with `--model` or `DECISION_MODEL` after probing a replacement model.

 Run the full evaluation and save a machine-readable report:

 ```powershell
 uv run python eval.py --output eval-report.json
 ```

 Use `--input 10k_financial_statements.json` for the app's JSON export, `--limit 1` for a low-cost smoke test, `--workers 4` to reduce API fan-out, or `--key-file path\to\api_key.txt` for another protected key location.

 ## SQLite tables

 - `filings`: one record per PDF and its full extracted text
 - `pages`: page-level source text for auditability
 - `metrics`: normalized line items, category, value, optional justification, and source page

 The statement dataset is curated from the statement pages in the attached filings. Important figures retain their source page in the app and in the exported reports.

## Accounting and performance report

`build_report.py` writes `reports/Alam_NVIDIA_Accounting_Performance.pdf` and the Excel backup `reports/NVIDIA_vs_AMD_Intel_DuPont_backup.xlsx`.

DuPont ratios, the peer comparison, and the residual-income forecast are quarterly. The ratio window is Q3 2024 through Q2 2026. Q2 2024 is kept only as the opening balance for averages. NVIDIA's quarter ends about four weeks after AMD and Intel, so the comparison pairs quarters by calendar slot. The Form 10-K discussion in the PDF is the strategic snapshot. NVIDIA FY t still sits next to AMD and Intel FY t−1 in that snapshot, because those are the 10-Ks.

Traditional DuPont is net margin times sales over average assets times average assets over average equity, on the quarter. Annualized rows are that quarter times four. Trailing twelve months is the sum of four quarters over the average of the opening and closing balance. Modified DuPont follows the course slide: NOPAT is net income plus after-tax net interest at the 21% statutory rate, and net debt is debt plus the long-term operating-lease line, minus cash and liquid investments. AMD's quarterly interest income is not disclosed, so modified DuPont is left blank for AMD. The Excel sheet computes the ratios and tests each identity. Averages are written as `(((opening)+(closing))/2)`.

`quarterly_statements.py` holds the quarter lines. `quarterly_model.py` computes the ratios, the trailing comparison, and the twenty-quarter forecast. `write_quarterly_pages.py` writes `dupont-analysis.html` and `forecast-valuation.html` from that model. `buy-sell-report.html` and `quick-pricing-formulas.html` still show the older annual model, and their notes say so.

```powershell
uv sync
uv run python build_report.py
uv run python write_quarterly_pages.py
```

The PDF is written to `reports/Alam_NVIDIA_Accounting_Performance.pdf`. Chart images are written to `reports/charts/` and are not source files.
