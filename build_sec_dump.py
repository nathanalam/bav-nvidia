"""Dump reported SEC facts for the quarterly balance sheet and income statement.

Values are the companyfacts figures. Concepts are not added together, blanks are
not filled, and a fourth quarter is not derived from the year. The income
statement includes the three-month fact and the six-month, nine-month, and
full-year facts, because the fourth quarter is often tagged only as a year.
"""

from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from urllib.request import Request, urlopen

import xlsxwriter

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "reports" / "SEC_quarterly_facts.xlsx"
CACHE = Path.home() / "AppData" / "Local" / "Temp" / "bav-facts"
WINDOW_START = "2023-07-01"
FORMS = {"10-Q", "10-Q/A", "10-K", "10-K/A"}
COMPANIES = (
    ("NVIDIA", "0001045810"),
    ("AMD", "0000002488"),
    ("Intel", "0000050863"),
)

# statement, concept, label. One concept, one row. Nothing is combined.
LINES = (
    ("Income statement", "Revenues", "Revenue"),
    ("Income statement", "RevenueFromContractWithCustomerExcludingAssessedTax", "Revenue"),
    ("Income statement", "CostOfRevenue", "Cost of revenue"),
    ("Income statement", "CostOfGoodsAndServicesSold", "Cost of sales"),
    ("Income statement", "GrossProfit", "Gross profit"),
    ("Income statement", "ResearchAndDevelopmentExpense", "Research and development"),
    ("Income statement", "SellingGeneralAndAdministrativeExpense", "Selling, general and administrative"),
    ("Income statement", "AmortizationOfIntangibleAssets", "Amortization of intangibles"),
    ("Income statement", "OperatingExpenses", "Operating expenses"),
    ("Income statement", "OperatingIncomeLoss", "Operating income"),
    ("Income statement", "InvestmentIncomeInterest", "Interest income"),
    ("Income statement", "InterestExpenseNonoperating", "Interest expense"),
    ("Income statement", "InterestExpense", "Interest expense"),
    ("Income statement", "GainLossOnInvestments", "Gain or loss on investments"),
    ("Income statement", "OtherNonoperatingIncomeExpense", "Other income (expense)"),
    ("Income statement", "NonoperatingIncomeExpense", "Nonoperating income (expense)"),
    ("Income statement", "IncomeLossFromEquityMethodInvestments", "Equity-method income"),
    ("Income statement", "IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest", "Income before tax"),
    ("Income statement", "IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments", "Income before tax and equity-method income"),
    ("Income statement", "IncomeTaxExpenseBenefit", "Income tax expense"),
    ("Income statement", "ProfitLoss", "Net income, consolidated"),
    ("Income statement", "NetIncomeLoss", "Net income"),
    ("Income statement", "NetIncomeLossAttributableToNoncontrollingInterest", "Net income attributable to noncontrolling interest"),
    ("Income statement", "EarningsPerShareBasic", "Basic earnings per share"),
    ("Income statement", "EarningsPerShareDiluted", "Diluted earnings per share"),
    ("Income statement", "WeightedAverageNumberOfSharesOutstandingBasic", "Basic weighted-average shares"),
    ("Income statement", "WeightedAverageNumberOfDilutedSharesOutstanding", "Diluted weighted-average shares"),
    ("Balance sheet", "CashAndCashEquivalentsAtCarryingValue", "Cash and cash equivalents"),
    ("Balance sheet", "CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents", "Cash, equivalents, and restricted cash"),
    ("Balance sheet", "MarketableSecuritiesCurrent", "Marketable securities, current"),
    ("Balance sheet", "DebtSecuritiesCurrent", "Debt securities, current"),
    ("Balance sheet", "EquitySecuritiesFvNi", "Equity securities at fair value"),
    ("Balance sheet", "ShortTermInvestments", "Short-term investments"),
    ("Balance sheet", "AvailableForSaleSecuritiesDebtSecuritiesCurrent", "Available-for-sale debt securities, current"),
    ("Balance sheet", "MarketableSecurities", "Marketable securities"),
    ("Balance sheet", "AccountsReceivableNetCurrent", "Accounts receivable"),
    ("Balance sheet", "InventoryNet", "Inventories"),
    ("Balance sheet", "PrepaidExpenseAndOtherAssetsCurrent", "Prepaid expenses and other current assets"),
    ("Balance sheet", "OtherAssetsCurrent", "Other current assets"),
    ("Balance sheet", "AssetsCurrent", "Total current assets"),
    ("Balance sheet", "PropertyPlantAndEquipmentNet", "Property, plant and equipment"),
    ("Balance sheet", "PropertyPlantAndEquipmentAndFinanceLeaseRightOfUseAssetAfterAccumulatedDepreciationAndAmortization", "Property, plant and equipment, including finance leases"),
    ("Balance sheet", "OperatingLeaseRightOfUseAsset", "Operating lease assets"),
    ("Balance sheet", "Goodwill", "Goodwill"),
    ("Balance sheet", "IntangibleAssetsNetExcludingGoodwill", "Intangible assets"),
    ("Balance sheet", "DeferredIncomeTaxAssetsNet", "Deferred tax assets"),
    ("Balance sheet", "EquitySecuritiesWithoutReadilyDeterminableFairValueAmount", "Equity securities without a readily determinable fair value"),
    ("Balance sheet", "OtherAssetsNoncurrent", "Other assets"),
    ("Balance sheet", "Assets", "Total assets"),
    ("Balance sheet", "AccountsPayableCurrent", "Accounts payable"),
    ("Balance sheet", "AccruedLiabilitiesCurrent", "Accrued liabilities"),
    ("Balance sheet", "EmployeeRelatedLiabilitiesCurrent", "Employee-related liabilities"),
    ("Balance sheet", "OtherLiabilitiesCurrent", "Other current liabilities"),
    ("Balance sheet", "DebtCurrent", "Debt, current"),
    ("Balance sheet", "LongTermDebtCurrent", "Long-term debt, current"),
    ("Balance sheet", "ShortTermBorrowings", "Short-term borrowings"),
    ("Balance sheet", "LiabilitiesCurrent", "Total current liabilities"),
    ("Balance sheet", "LongTermDebt", "Long-term debt, including current"),
    ("Balance sheet", "LongTermDebtNoncurrent", "Long-term debt"),
    ("Balance sheet", "OperatingLeaseLiabilityNoncurrent", "Operating lease liabilities"),
    ("Balance sheet", "OtherLiabilitiesNoncurrent", "Other long-term liabilities"),
    ("Balance sheet", "Liabilities", "Total liabilities"),
    ("Balance sheet", "CommonStockValue", "Common stock"),
    ("Balance sheet", "AdditionalPaidInCapital", "Additional paid-in capital"),
    ("Balance sheet", "AdditionalPaidInCapitalCommonStock", "Additional paid-in capital and common stock"),
    ("Balance sheet", "CommonStocksIncludingAdditionalPaidInCapital", "Common stock and additional paid-in capital"),
    ("Balance sheet", "AccumulatedOtherComprehensiveIncomeLossNetOfTax", "Accumulated other comprehensive income"),
    ("Balance sheet", "RetainedEarningsAccumulatedDeficit", "Retained earnings"),
    ("Balance sheet", "TreasuryStockValue", "Treasury stock"),
    ("Balance sheet", "StockholdersEquity", "Stockholders' equity"),
    ("Balance sheet", "MinorityInterest", "Noncontrolling interest"),
    ("Balance sheet", "StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest", "Equity including noncontrolling interest"),
    ("Balance sheet", "LiabilitiesAndStockholdersEquity", "Total liabilities and equity"),
)


def months_of(start: str | None, end: str) -> int | None:
    if not start:
        return None
    span = (datetime.fromisoformat(end) - datetime.fromisoformat(start)).days
    if 75 <= span <= 105:
        return 3
    if 160 <= span <= 200:
        return 6
    if 250 <= span <= 290:
        return 9
    if 340 <= span <= 380:
        return 12
    return None


def companyfacts(cik: str) -> dict:
    CACHE.mkdir(parents=True, exist_ok=True)
    path = CACHE / f"{cik}.json"
    if not path.exists():
        request = Request(
            f"https://data.sec.gov/api/xbrl/companyfacts/CIK{cik}.json",
            headers={"User-Agent": "bav-nvidia research nathan@example.com"},
        )
        path.write_bytes(urlopen(request, timeout=120).read())
    return json.loads(path.read_text(encoding="utf-8"))


def collect() -> list[dict]:
    catalog = {(statement, concept): label for statement, concept, label in LINES}
    order = {(statement, concept): index for index, (statement, concept, _) in enumerate(LINES)}
    rows = []
    for company, cik in COMPANIES:
        gaap = companyfacts(cik)["facts"].get("us-gaap", {})
        for (statement, concept), label in catalog.items():
            body = gaap.get(concept)
            if not body:
                continue
            best: dict[tuple, dict] = {}
            for unit, facts in body.get("units", {}).items():
                for fact in facts:
                    if fact.get("form") not in FORMS or fact.get("end", "") < WINDOW_START:
                        continue
                    start = fact.get("start")
                    if statement == "Balance sheet":
                        if start:
                            continue
                        months = None
                    else:
                        months = months_of(start, fact["end"])
                        if months is None:
                            continue
                    key = (unit, start or "", fact["end"], months)
                    current = best.get(key)
                    if current is None or (fact.get("filed", ""), fact.get("accn", "")) > (
                        current.get("filed", ""),
                        current.get("accn", ""),
                    ):
                        best[key] = fact
            for (unit, start, end, months), fact in best.items():
                rows.append(
                    {
                        "company": company,
                        "cik": cik,
                        "statement": statement,
                        "line": label,
                        "concept": concept,
                        "unit": unit,
                        "period_start": start,
                        "period_end": end,
                        "months": months,
                        "fy": fact.get("fy"),
                        "fp": fact.get("fp"),
                        "form": fact.get("form"),
                        "filed": fact.get("filed"),
                        "accession": fact.get("accn"),
                        "value": fact["val"],
                        "_order": order[(statement, concept)],
                    }
                )
    rows.sort(key=lambda row: (row["company"], row["statement"], row["_order"], row["period_end"], row["months"] or 0))
    return rows


def write(rows: list[dict]) -> None:
    OUT.parent.mkdir(parents=True, exist_ok=True)
    workbook = xlsxwriter.Workbook(str(OUT))
    header = workbook.add_format({"bold": True, "bg_color": "#1B2A4A", "font_color": "white", "font_name": "Calibri"})
    text = workbook.add_format({"font_name": "Calibri"})
    number = workbook.add_format({"font_name": "Calibri", "num_format": "#,##0.00"})
    note = workbook.add_format({"font_name": "Calibri", "text_wrap": True, "valign": "top"})
    facts = workbook.add_worksheet("Facts")
    columns = [
        "company", "cik", "statement", "line", "concept", "unit",
        "period_start", "period_end", "months", "fy", "fp", "form", "filed", "accession", "value",
    ]
    for col, name in enumerate(columns):
        facts.write(0, col, name, header)
    for index, row in enumerate(rows, start=1):
        for col, name in enumerate(columns):
            value = row[name]
            if name == "value":
                facts.write_number(index, col, value, number)
            elif value is None:
                facts.write_blank(index, col, None, text)
            else:
                facts.write(index, col, value, text)
    facts.autofilter(0, 0, len(rows), len(columns) - 1)
    facts.freeze_panes(1, 0)
    widths = [14, 14, 20, 62, 78, 14, 14, 14, 10, 8, 8, 10, 12, 24, 22]
    for col, width in enumerate(widths):
        facts.set_column(col, col, width)

    readme = workbook.add_worksheet("Readme")
    readme.set_column(0, 0, 120)
    readme.write(0, 0, "SEC quarterly facts, as reported", header)
    notes = [
        "Source: SEC companyfacts, us-gaap, Form 10-Q and Form 10-K, including amendments. The companyfacts feed has no segment breakdowns.",
        "Companies: NVIDIA (0001045810), AMD (0000002488), Intel (0000050863).",
        "Window: period end on or after 2023-07-01, through the latest quarter in the downloaded facts.",
        "value is the figure reported to the SEC. Dollars are dollars, shares are shares, and per-share amounts are dollars per share. Nothing is summed, scaled, or filled with zero.",
        "months is 3, 6, 9, or 12 for an income-statement period. Balance-sheet rows are the instant at period_end, so months is blank.",
        "A line missing from a quarter was not tagged for that period. Fourth-quarter income-statement amounts are often absent as a three-month fact. The nine-month and full-year facts are both here.",
        "Where the same period was filed more than once, the latest filing is the one kept. fy and fp are the tags in that filing.",
        "Two concepts can describe the same caption. They are separate rows. Debt securities and equity securities are not added together, and interest income is not inferred.",
    ]
    for index, line in enumerate(notes, start=2):
        readme.write(index, 0, line, note)
        readme.set_row(index, 32)
    workbook.close()


def main() -> None:
    rows = collect()
    anchors = {
        ("NVIDIA", "Revenues", "2026-07-26", 3): 96_221_000_000,
        ("AMD", "RevenueFromContractWithCustomerExcludingAssessedTax", "2026-06-27", 3): 11_536_000_000,
        ("Intel", "ProfitLoss", "2026-06-27", 3): -10_848_000_000,
        ("NVIDIA", "Assets", "2026-07-26", None): 320_272_000_000,
        ("AMD", "Assets", "2026-06-27", None): 84_464_000_000,
        ("Intel", "Assets", "2026-06-27", None): 202_439_000_000,
    }
    for (company, concept, end, months), expected in anchors.items():
        found = [
            row["value"]
            for row in rows
            if row["company"] == company and row["concept"] == concept and row["period_end"] == end and row["months"] == months
        ]
        if found != [expected]:
            raise AssertionError((company, concept, end, months, found))
    amd_interest = [
        row for row in rows
        if row["company"] == "AMD" and row["concept"] == "InvestmentIncomeInterest" and row["months"] == 3
    ]
    if amd_interest:
        raise AssertionError("AMD quarterly interest income was not expected as a three-month fact")
    write(rows)
    print(f"wrote {OUT}")
    print(f"rows {len(rows)}")
    for company, _cik in COMPANIES:
        ends = sorted({row["period_end"] for row in rows if row["company"] == company and (row["months"] == 3 or row["statement"] == "Balance sheet")})
        print(company, "period ends", len(ends), ends[0], ends[-1])


if __name__ == "__main__":
    main()
