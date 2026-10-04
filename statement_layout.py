"""Shared presentation rules for reconstructed financial statements."""

from __future__ import annotations

import pandas as pd

UNIT_CAPTION = "Amounts in USD millions, except per-share data and share counts"

STATEMENT_ORDER = {
    "Balance sheet": ["Cash and cash equivalents", "Marketable securities", "Short-term investments", "Accounts receivable", "Inventories", "Prepaid expenses and other current assets", "Total current assets", "Property, plant and equipment", "Operating lease right-of-use assets", "Operating lease assets", "Goodwill", "Intangible assets", "Deferred income tax assets", "Other assets", "Total assets", "Accounts payable", "Accrued and other current liabilities", "Other current liabilities", "Current portion of long-term debt", "Short-term debt", "Total current liabilities", "Long-term debt", "Other long-term liabilities", "Total liabilities", "Preferred stock", "Common stock", "Additional paid-in capital", "Treasury stock", "Accumulated other comprehensive income", "Accumulated deficit", "Retained earnings", "Stockholders' equity", "Total liabilities and equity"],
    "Income statement": ["Revenue", "Cost of sales", "Cost of revenue", "Total cost of sales", "Gross profit", "Research and development", "R&D", "Marketing, general and administrative", "Selling, general and administrative", "Amortization of acquisition-related intangibles", "Depreciation and amortization", "Licensing gain", "Total operating expenses", "Operating income", "Interest income", "Interest expense", "Other income (expense), net", "Income before income taxes and equity income", "Income before income taxes", "Income tax expense", "Equity income in investee", "Net income", "Basic EPS", "Diluted EPS", "Basic weighted average shares", "Diluted weighted average shares", "Weighted average shares"],
    "Cash flow statement": ["Net income", "Depreciation and amortization", "Depreciation", "Stock-based compensation", "Deferred income taxes", "Accounts receivable", "Inventories", "Accounts payable", "Other assets", "Other liabilities", "Operating cash flow", "Capital expenditures", "Investing cash flow", "Cash at beginning of period", "Net change in cash", "Cash at end of period", "Financing cash flow"],
}

TOTAL_LINE_ITEMS = {"Total current assets", "Total assets", "Total current liabilities", "Total liabilities", "Total liabilities and equity", "Stockholders' equity", "Total cost of sales", "Gross profit", "Total operating expenses", "Operating income", "Income before income taxes", "Net income", "Operating cash flow", "Investing cash flow", "Financing cash flow", "Net change in cash", "Cash at end of period"}

def order_statement_frame(frame: pd.DataFrame, category: str) -> pd.DataFrame:
    if frame.empty or "Line item" not in frame.columns:
        return frame
    preferred = {label: index for index, label in enumerate(STATEMENT_ORDER.get(category, []))}
    result = frame.copy()
    result["_statement_order"] = result["Line item"].map(preferred).fillna(10000)
    result["_original_order"] = range(len(result))
    result = result.sort_values(["_statement_order", "_original_order"], kind="stable")
    return result.drop(columns=["_statement_order", "_original_order"])
