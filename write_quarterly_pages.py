"""Write the quarterly DuPont and forecast pages from the calculation model.

The pages are dupont-analysis.html and forecast-valuation.html. Both read
RATIOS, the trailing windows, and the twenty-quarter forecast. They do not
keep a second set of annual ratios.
"""

from __future__ import annotations

from html import escape
from pathlib import Path

from quarterly_model import (
    FORECAST_ASSUMPTIONS,
    FORECASTS,
    ORDER,
    RATIOS,
    SLOTS,
    forecast,
    latest_ttm,
    ttm_at,
    year_totals,
)
from quarterly_statements import QUARTERS

ROOT = Path(__file__).resolve().parent
MARKET = {"NVIDIA": 236.16, "AMD": 631.57, "Intel": 117.40}
PRIOR = {"NVIDIA": 54.06, "AMD": 44.38, "Intel": 8.03}
KLASS = {"NVIDIA": "nvidia", "AMD": "amd", "Intel": "intel"}

CSS = """
    :root { color-scheme: light dark; }
    * { box-sizing: border-box; }
    body {
      font-family: -apple-system, system-ui, 'Segoe UI', sans-serif;
      max-width: 1100px; width: 100%; margin: 0 auto;
      padding: clamp(16px, 4vw, 32px);
      color: light-dark(#1e293b, #e2e8f0);
      background: light-dark(#ffffff, #0f172a);
      line-height: 1.6;
    }
    h1 { font-size: 1.75rem; margin-bottom: 0.25em; color: light-dark(#0f172a, #f1f5f9); }
    h2 { font-size: 1.35rem; margin-top: 2em; color: light-dark(#1e40af, #93c5fd); border-bottom: 2px solid light-dark(#dbeafe, #1e3a5f); padding-bottom: 0.3em; }
    h3 { font-size: 1.1rem; margin-top: 1.5em; color: light-dark(#374151, #d1d5db); }
    p { margin: 0.6em 0; }
    a { color: light-dark(#1d4ed8, #93c5fd); }
    .subtitle { color: light-dark(#64748b, #94a3b8); margin-top: 0; font-size: 0.95rem; }
    .formula { background: light-dark(#f0f9ff, #1e293b); border: 1px solid light-dark(#bfdbfe, #334155); border-radius: 6px; padding: 10px 16px; margin: 12px 0; font-family: 'SF Mono', 'Fira Code', monospace; font-size: 0.88rem; }
    .note { background: light-dark(#fffbeb, #1c1917); border-left: 3px solid light-dark(#f59e0b, #d97706); padding: 8px 14px; margin: 12px 0; font-size: 0.85rem; color: light-dark(#92400e, #fbbf24); }
    .info { background: light-dark(#eff6ff, #0c1929); border-left: 3px solid light-dark(#3b82f6, #60a5fa); padding: 8px 14px; margin: 12px 0; font-size: 0.85rem; color: light-dark(#1e40af, #93c5fd); }
    .warning { background: light-dark(#fef2f2, #1c0a0a); border-left: 3px solid light-dark(#ef4444, #dc2626); padding: 8px 14px; margin: 12px 0; font-size: 0.85rem; color: light-dark(#991b1b, #fca5a5); }
    .table-wrap { overflow-x: auto; margin: 16px 0; }
    table { border-collapse: collapse; width: 100%; font-size: 0.82rem; }
    th, td { border: 1px solid light-dark(#e2e8f0, #334155); padding: 5px 8px; text-align: right; white-space: nowrap; }
    th { background: light-dark(#f1f5f9, #1e293b); font-weight: 600; text-align: center; }
    th:first-child, td:first-child { text-align: left; }
    .section td { background: light-dark(#dbeafe, #1e3a5f); font-weight: 700; text-align: left; }
    .positive { color: light-dark(#15803d, #4ade80); }
    .negative { color: light-dark(#dc2626, #f87171); }
    .bold { font-weight: 700; }
    .insight-box { background: light-dark(#f0fdf4, #052e16); border: 1px solid light-dark(#86efac, #166534); border-radius: 6px; padding: 12px 16px; margin: 16px 0; }
    .insight-box h4 { margin: 0 0 6px 0; color: light-dark(#166534, #4ade80); font-size: 0.95rem; }
    .insight-box ul { margin: 4px 0; padding-left: 20px; font-size: 0.88rem; }
    .kpi-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(140px, 1fr)); gap: 12px; margin: 12px 0 20px; }
    .kpi { background: light-dark(#f8fafc, #1e293b); border: 1px solid light-dark(#e2e8f0, #334155); border-radius: 8px; padding: 12px; text-align: center; }
    .kpi-label { font-size: 0.72rem; color: light-dark(#64748b, #94a3b8); text-transform: uppercase; letter-spacing: 0.05em; }
    .kpi-value { font-size: 1.35rem; font-weight: 700; margin-top: 2px; }
    .kpi-sub { font-size: 0.75rem; color: light-dark(#94a3b8, #64748b); margin-top: 2px; }
    .nvidia { color: light-dark(#16a34a, #4ade80); }
    .amd { color: light-dark(#dc2626, #f87171); }
    .intel { color: light-dark(#2563eb, #60a5fa); }
    .cards { display: grid; grid-template-columns: repeat(auto-fit, minmax(260px, 1fr)); gap: 16px; }
    .card { background: light-dark(#f8fafc, #1e293b); border: 1px solid light-dark(#e2e8f0, #334155); border-radius: 8px; padding: 14px; }
    .card h3 { margin-top: 0; }
    .latest { box-shadow: inset 0 -2px 0 light-dark(#16a34a, #4ade80); }
"""


def pct(value: float | None, digits: int = 1) -> str:
    if value is None:
        return "n.m."
    return f"{value * 100:.{digits}f}%"


def multiple(value: float | None, digits: int = 2) -> str:
    if value is None:
        return "n.m."
    return f"{value:.{digits}f}x"


def billions(value: float | None) -> str:
    if value is None:
        return "n.m."
    sign = "-" if value < 0 else ""
    return f"{sign}${abs(value) / 1000:.1f}bn"


def millions(value: float | None) -> str:
    if value is None:
        return "n.m."
    sign = "-" if value < 0 else ""
    return f"{sign}{abs(value):,.0f}"


def dollars(value: float) -> str:
    return f"${value:,.2f}"


def net_position(value: float) -> str:
    if value < 0:
        return f"Net cash {billions(abs(value))}"
    return f"Net debt {billions(value)}"


def page(title: str, body: str) -> str:
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>{escape(title)}</title>
  <style>{CSS}</style>
</head>
<body>
{body}
</body>
</html>
"""


def cell(text: str, *, negative: bool = False, bold: bool = False, header: bool = False) -> str:
    classes = []
    if negative:
        classes.append("negative")
    if bold:
        classes.append("bold")
    attr = f' class="{" ".join(classes)}"' if classes else ""
    tag = "th" if header else "td"
    return f"<{tag}{attr}>{text}</{tag}>"


def num_cell(value: float | None, text: str, *, bold: bool = False) -> str:
    return cell(text, negative=value is not None and value < 0, bold=bold)


def wrap_table(rows: list[str]) -> str:
    return '<div class="table-wrap"><table>\n' + "\n".join(rows) + "\n</table></div>"


def section_row(label: str, span: int) -> str:
    return f'<tr class="section"><td colspan="{span}">{escape(label)}</td></tr>'


def statement(name: str, slot_index: int) -> dict:
    return QUARTERS[name][slot_index + 1]


def ratio_table(name: str) -> str:
    ratios = RATIOS[name]
    span = 1 + len(SLOTS)
    headers = ["<th>Line</th>"]
    dates = ["<th>Quarter-end</th>"]
    for index, slot in enumerate(SLOTS):
        klass = ' class="latest"' if index == len(SLOTS) - 1 else ""
        headers.append(f"<th{klass}>{escape(slot)}</th>")
        dates.append(f"<th{klass}>{escape(statement(name, index)['end'])}</th>")
    rows = ["<thead>", "<tr>" + "".join(headers) + "</tr>", "<tr>" + "".join(dates) + "</tr>", "</thead>", "<tbody>"]

    def add(label: str, values: list[str], *, bold: bool = False) -> None:
        body = [cell(escape(label), bold=bold)]
        body.extend(values)
        rows.append("<tr>" + "".join(body) + "</tr>")

    def from_statement(key: str) -> None:
        values = []
        for index in range(len(SLOTS)):
            amount = statement(name, index)[key]
            values.append(num_cell(amount, millions(amount)))
        add(key_labels[key], values)

    def from_ratio(label: str, key: str, kind: str, *, bold: bool = False, digits: int = 2) -> None:
        values = []
        for row in ratios:
            value = row[key]
            if kind == "pct":
                text = pct(value, digits)
            else:
                text = multiple(value, digits)
            values.append(num_cell(value, text, bold=bold))
        add(label, values, bold=bold)

    rows.append(section_row("Statement lines, $ millions", span))
    for key in ("rev", "oi", "ni"):
        from_statement(key)
    rows.append(section_row("Traditional DuPont, on the quarter", span))
    from_ratio("Net margin", "npm", "pct")
    from_ratio("Asset turnover", "ato", "x")
    from_ratio("Equity multiplier", "em", "x")
    from_ratio("ROE", "roe", "pct", bold=True)
    add("Identity: margin x turnover x leverage", [cell("YES", bold=True) for _ in SLOTS], bold=True)
    rows.append(section_row("Modified DuPont, on the quarter", span))
    from_ratio("NOPAT margin", "pm", "pct")
    from_ratio("NOA turnover", "turn", "x")
    from_ratio("Operating ROA", "rnoa", "pct", bold=True)
    from_ratio("Net borrowing cost", "nbc", "pct")
    from_ratio("Spread", "spread", "pct")
    from_ratio("Net financial leverage", "flev", "x")
    from_ratio("Leverage gain", "gain", "pct", bold=True)
    from_ratio("ROE", "roe", "pct", bold=True)
    checks = []
    for row in ratios:
        checks.append(cell("n.m." if row["rnoa"] is None else "YES", bold=True))
    add("Identity: operating ROA + gain", checks, bold=True)
    from_ratio("Operating income x 0.79 / average NOA", "alt", "pct")
    rows.append(section_row("Annualized, the quarter times four. Margins are not annualized.", span))
    from_ratio("ROE, annualized", "roe_ann", "pct", bold=True)
    from_ratio("Asset turnover, annualized", "ato_ann", "x")
    from_ratio("Operating ROA, annualized", "rnoa_ann", "pct", bold=True)
    from_ratio("Taxed operating income / average NOA, annualized", "alt_ann", "pct")
    rows.append("</tbody>")
    return wrap_table(rows)


key_labels = {
    "rev": "Revenue",
    "oi": "Operating income",
    "ni": "Net income",
}


def comparison_table() -> str:
    headers = ["<th>Line</th>"] + [f'<th class="{KLASS[name]}">{escape(name)}</th>' for name in ORDER]
    rows = ["<thead><tr>" + "".join(headers) + "</tr></thead><tbody>"]

    def add(label: str, texts: list[str], *, bold: bool = False) -> None:
        body = [cell(escape(label), bold=bold)] + texts
        rows.append("<tr>" + "".join(body) + "</tr>")

    latest = {name: RATIOS[name][-1] for name in ORDER}
    trailing = {name: latest_ttm(name) for name in ORDER}
    rows.append(section_row("Latest quarter", 4))
    add("Quarter", [cell(escape(SLOTS[-1])) for _ in ORDER])
    add("Quarter-end", [cell(escape(statement(name, len(SLOTS) - 1)["end"])) for name in ORDER])
    add("Revenue", [num_cell(statement(name, len(SLOTS) - 1)["rev"], billions(statement(name, len(SLOTS) - 1)["rev"])) for name in ORDER])
    add("Operating margin", [num_cell(latest[name]["opm"], pct(latest[name]["opm"])) for name in ORDER])
    add("ROE, annualized", [num_cell(latest[name]["roe_ann"], pct(latest[name]["roe_ann"]), bold=True) for name in ORDER], bold=True)
    add("Operating ROA, annualized", [num_cell(latest[name]["rnoa_ann"], pct(latest[name]["rnoa_ann"])) for name in ORDER])
    add("Taxed operating ROA, annualized", [num_cell(latest[name]["alt_ann"], pct(latest[name]["alt_ann"])) for name in ORDER])
    add("Balance sheet", [cell(escape(net_position(latest[name]["nd"]))) for name in ORDER])
    rows.append(section_row("Trailing twelve months, Q3 2025 through Q2 2026", 4))
    add("Revenue", [num_cell(trailing[name]["rev"], billions(trailing[name]["rev"])) for name in ORDER])
    add("Net income", [num_cell(trailing[name]["ni"], billions(trailing[name]["ni"]), bold=True) for name in ORDER], bold=True)
    add("Operating margin", [num_cell(trailing[name]["opm"], pct(trailing[name]["opm"])) for name in ORDER])
    add("Net margin", [num_cell(trailing[name]["npm"], pct(trailing[name]["npm"])) for name in ORDER])
    add("ROE", [num_cell(trailing[name]["roe"], pct(trailing[name]["roe"]), bold=True) for name in ORDER], bold=True)
    add("Asset turnover", [num_cell(trailing[name]["ato"], multiple(trailing[name]["ato"])) for name in ORDER])
    add("Operating ROA", [num_cell(trailing[name]["rnoa"], pct(trailing[name]["rnoa"])) for name in ORDER])
    add("Taxed operating ROA", [num_cell(trailing[name]["alt"], pct(trailing[name]["alt"])) for name in ORDER])
    add("Identity", [cell("YES", bold=True) for _ in ORDER], bold=True)
    rows.append(section_row("Trailing ROE by window end", 4))
    for index in range(3, len(SLOTS)):
        window = {name: ttm_at(name, index) for name in ORDER}
        add(SLOTS[index], [num_cell(window[name]["roe"], pct(window[name]["roe"])) for name in ORDER])
    rows.append("</tbody>")
    return wrap_table(rows)


def dupont_page() -> str:
    latest = {name: RATIOS[name][-1] for name in ORDER}
    trailing = {name: latest_ttm(name) for name in ORDER}
    cards = []
    for name in ORDER:
        row = latest[name]
        trail = trailing[name]
        cards.append(f"""
    <div class="card">
      <h3 class="{KLASS[name]}">{escape(name)}</h3>
      <div class="kpi-grid">
        <div class="kpi"><div class="kpi-label">Latest operating margin</div><div class="kpi-value {KLASS[name]}">{pct(row['opm'])}</div><div class="kpi-sub">{escape(SLOTS[-1])} · {escape(statement(name, len(SLOTS) - 1)['end'])}</div></div>
        <div class="kpi"><div class="kpi-label">ROE, annualized</div><div class="kpi-value {KLASS[name]}">{pct(row['roe_ann'])}</div><div class="kpi-sub">Quarter x 4 · quarter itself {pct(row['roe'], 2)}</div></div>
        <div class="kpi"><div class="kpi-label">Trailing ROE</div><div class="kpi-value {KLASS[name]}">{pct(trail['roe'])}</div><div class="kpi-sub">Net income {billions(trail['ni'])}</div></div>
        <div class="kpi"><div class="kpi-label">Trailing operating margin</div><div class="kpi-value {KLASS[name]}">{pct(trail['opm'])}</div><div class="kpi-sub">Revenue {billions(trail['rev'])}</div></div>
      </div>
    </div>""")
    company_blocks = []
    for name in ORDER:
        company_blocks.append(f'<h3 class="{KLASS[name]}">{escape(name)}</h3>\n{ratio_table(name)}')
    nvidia = latest["NVIDIA"]
    amd = latest["AMD"]
    intel = latest["Intel"]
    body = f"""
<h1>Quarterly DuPont analysis</h1>
<p class="subtitle">NVIDIA, AMD, and Intel · Q3 2024 through Q2 2026 · Prepared October 6, 2026</p>
<div class="note">
  Ratios, the comparison, and the forecast use quarterly statements. The Form 10-K remains the strategic source in the PDF. Q2 2024 is the opening balance for the averages and is not a ratio column. NVIDIA's quarter ends about four weeks after AMD and Intel, so the columns are calendar slots.
</div>
<div class="cards">
{''.join(cards)}
</div>
<h2>How the identity is computed</h2>
<div class="formula">Quarterly ROE = net margin x asset turnover x equity multiplier = (net income / revenue) x (revenue / average assets) x (average assets / average equity)</div>
<div class="formula">Modified ROE = operating ROA + (operating ROA − net borrowing cost) x net financial leverage</div>
<p>NOPAT is net income plus after-tax net interest at the 21% statutory rate. Net debt is debt plus the long-term operating-lease line, minus cash and liquid investments. Average balances are the opening quarter and the closing quarter. YES is the unrounded quarterly identity. An annualized row is that quarter times four, so it can sit next to an annual cost of equity. It is not a second identity. Trailing twelve months sums four quarters and divides by the average of the balance before the window and the balance at the end.</p>
<div class="info">AMD interest income is annual-only in the filing taxonomy, so modified DuPont is n.m. for AMD. The taxed-operating-income row is filled in for all three. Intel net income is consolidated profit, and Intel equity includes non-controlling interests.</div>
<h2>Eight ratio quarters</h2>
{''.join(company_blocks)}
<h2>Comparison</h2>
<p>The first block is the latest quarter. The second is the trailing twelve months ending Q2 2026, the same window as the buy-and-sell note's trailing net income. The identity row is net margin times turnover times leverage on those totals.</p>
{comparison_table()}
<h2>What the quarters say</h2>
<div class="insight-box">
  <h4>The run-rate is the trailing year, read next to the latest quarter</h4>
  <ul>
    <li>NVIDIA trailing revenue is {billions(trailing['NVIDIA']['rev'])} and trailing net income is {billions(trailing['NVIDIA']['ni'])}. Trailing ROE is {pct(trailing['NVIDIA']['roe'])}. The latest quarter annualizes to {pct(nvidia['roe_ann'])}. Annualized operating ROA is {pct(nvidia['rnoa_ann'])}. NVIDIA holds {net_position(nvidia['nd']).replace('Net cash ', 'net cash of ').replace('Net debt ', 'net debt of ')}.</li>
    <li>AMD trailing ROE is {pct(trailing['AMD']['roe'])} on revenue of {billions(trailing['AMD']['rev'])}. Latest-quarter annualized turnover is {multiple(amd['ato_ann'])}. Modified DuPont stays open. Taxed operating income over average net operating assets annualizes to {pct(amd['alt_ann'])}.</li>
    <li>Intel trailing operating margin is {pct(trailing['Intel']['opm'])} and trailing ROE is {pct(trailing['Intel']['roe'])}, on revenue of {billions(trailing['Intel']['rev'])}. The latest quarter's operating margin is {pct(intel['opm'])}, on operating income of {billions(statement('Intel', len(SLOTS) - 1)['oi'])} and consolidated net income of {billions(statement('Intel', len(SLOTS) - 1)['ni'])}. Intel carries {net_position(intel['nd']).replace('Net debt ', 'net debt of ').replace('Net cash ', 'net cash of ')}.</li>
    <li>Against the year-ago quarter, revenue is {pct(statement('NVIDIA', 7)['rev'] / statement('NVIDIA', 3)['rev'] - 1)} at NVIDIA, {pct(statement('AMD', 7)['rev'] / statement('AMD', 3)['rev'] - 1)} at AMD, and {pct(statement('Intel', 7)['rev'] / statement('Intel', 3)['rev'] - 1)} at Intel.</li>
  </ul>
</div>
<h2>Sources</h2>
<p>Lines are SEC companyfacts, USD millions, from the Form 10-Q and Form 10-K quarters. A tagged quarter is kept. The fourth quarter is the rounded fiscal year minus the three rounded quarters, so four quarters add to the 10-K. The plug is within $2 million of the unrounded residual. NVIDIA liquid investments are marketable securities through October 2025, and current debt securities plus equity securities at fair value from January 25, 2026. Intel face cash was not tagged on June 29, 2024, September 28, 2024, and March 29, 2025, so those three balances are cash plus restricted cash.</p>
<p>The forecast that uses these ratios is <a href="forecast-valuation.html">forecast-valuation.html</a>. The PDF and the Excel identities are written by <code>build_report.py</code>.</p>
<p class="subtitle">Analysis for study purposes, not investment advice.</p>
"""
    return page("Quarterly DuPont: NVIDIA, AMD, Intel", body)


def year_table(caption: str, key: str, kind: str) -> str:
    headers = ["<th></th>"] + [f'<th class="{KLASS[name]}">{escape(name)}</th>' for name in ORDER]
    rows = ["<thead><tr>" + "".join(headers) + "</tr></thead><tbody>"]
    for index in range(5):
        texts = []
        for name in ORDER:
            value = year_totals(name)[index][key]
            text = billions(value) if kind == "bn" else pct(value)
            texts.append(num_cell(value, text))
        rows.append("<tr>" + cell(f"{caption} {index + 1}") + "".join(texts) + "</tr>")
    if key == "rev":
        texts = []
        for name in ORDER:
            growth = year_totals(name)[0]["rev"] / latest_ttm(name)["rev"] - 1
            texts.append(num_cell(growth, pct(growth), bold=True))
        rows.append("<tr>" + cell("Year-1 revenue vs trailing", bold=True) + "".join(texts) + "</tr>")
    rows.append("</tbody>")
    return wrap_table(rows)


def bridge_table(sensitivity: dict) -> str:
    headers = ["<th>Bridge</th>"] + [f'<th class="{KLASS[name]}">{escape(name)}</th>' for name in ORDER]
    rows = ["<thead><tr>" + "".join(headers) + "</tr></thead><tbody>"]

    def add(label: str, texts: list[str], *, bold: bool = False) -> None:
        rows.append("<tr>" + cell(label, bold=bold) + "".join(texts) + "</tr>")

    add("Cost of equity", [cell(pct(FORECASTS[name]["ke"])) for name in ORDER])
    add("Terminal growth", [cell(pct(FORECASTS[name]["g"])) for name in ORDER])
    add("Opening book equity", [cell(billions(FORECASTS[name]["book"])) for name in ORDER])
    add("PV of residual income", [num_cell(FORECASTS[name]["pv_ri"], billions(FORECASTS[name]["pv_ri"])) for name in ORDER])
    add("PV of terminal value", [num_cell(FORECASTS[name]["pv_tv"], billions(FORECASTS[name]["pv_tv"])) for name in ORDER])
    add("Equity value", [num_cell(FORECASTS[name]["mve"], billions(FORECASTS[name]["mve"]), bold=True) for name in ORDER], bold=True)
    add("Diluted shares, millions", [cell(f"{FORECASTS[name]['shares']:,.0f}") for name in ORDER])
    add("Value per share", [cell(dollars(FORECASTS[name]["price"]), bold=True) for name in ORDER], bold=True)
    add("Prior annual model, per share", [cell(dollars(PRIOR[name])) for name in ORDER])
    add("Market price, Oct 5, 2026", [cell(dollars(MARKET[name])) for name in ORDER])
    add("Intel sensitivity, per share", [cell("—"), cell("—"), cell(dollars(sensitivity["price"]), bold=True)])
    rows.append("</tbody>")
    return wrap_table(rows)


def quarter_table(name: str) -> str:
    forecast_row = FORECASTS[name]
    ke_q = forecast_row["ke"] / 4
    headers = ["<th>Q</th>", "<th>Revenue, $m</th>", "<th>Operating margin</th>", "<th>Operating income, $m</th>", "<th>Net income, $m</th>", "<th>Equity, end, $m</th>", "<th>Residual income, $m</th>", "<th>Present value, $m</th>"]
    rows = ["<thead><tr>" + "".join(headers) + "</tr></thead><tbody>"]
    for item in forecast_row["quarters"]:
        present = item["residual"] / (1 + ke_q) ** item["t"]
        values = [
            cell(str(item["t"])),
            num_cell(item["rev"], millions(item["rev"])),
            num_cell(item["om"], pct(item["om"])),
            num_cell(item["oi"], millions(item["oi"])),
            num_cell(item["ni"], millions(item["ni"])),
            num_cell(item["eq"], millions(item["eq"])),
            num_cell(item["residual"], millions(item["residual"])),
            num_cell(present, millions(present)),
        ]
        rows.append("<tr>" + "".join(values) + "</tr>")
    rows.append("</tbody>")
    return wrap_table(rows)


def forecast_page(sensitivity: dict) -> str:
    assumption_cards = []
    for name in ORDER:
        assumptions = FORECAST_ASSUMPTIONS[name]
        path = FORECASTS[name]
        assumption_cards.append(f"""
    <div class="card">
      <h3 class="{KLASS[name]}">{escape(name)}</h3>
      <p>Cost of equity {pct(assumptions['ke'])}. Terminal growth {pct(assumptions['g'])}. Operating margin glides from the trailing {pct(path['om0'])} to {pct(assumptions['om_target'])}. Net income is {assumptions['ni_on_oi']:.4f} times operating income. Turnover stays at {multiple(path['ato'])}. Equity stays at {pct(path['eq_ratio'])} of assets.</p>
      <p>{escape(assumptions['note'])}</p>
    </div>""")
    growth = {name: year_totals(name)[0]["rev"] / latest_ttm(name)["rev"] - 1 for name in ORDER}
    detail = []
    for name in ORDER:
        detail.append(f'<h3 class="{KLASS[name]}">{escape(name)}</h3>\n{quarter_table(name)}')
    body = f"""
<h1>Twenty-quarter residual-income forecast</h1>
<p class="subtitle">NVIDIA, AMD, and Intel · off the latest quarter · Prepared October 6, 2026</p>
<div class="warning"><strong>Not investment advice.</strong> One explicit path. The market prices below are the October 5, 2026 snapshot in the buy-and-sell note. This page does not update those prices and does not change that note's verdicts.</div>
<div class="info">Each of the twenty quarters is its own residual-income period. The earlier quarterly prices on this page were a straight line between annual values. This page replaces that line. Cost of equity stays at 13.5% for NVIDIA, 12.5% for AMD, and 9.7% for Intel. Terminal growth stays at 3% a year. The scenario rates in <a href="buy-sell-report.html">buy-sell-report.html</a> stay on that page.</div>
<h2>Fair value</h2>
<div class="kpi-grid">
  <div class="kpi"><div class="kpi-label">NVIDIA</div><div class="kpi-value nvidia">{dollars(FORECASTS['NVIDIA']['price'])}</div><div class="kpi-sub">Annual model {dollars(PRIOR['NVIDIA'])} · market {dollars(MARKET['NVIDIA'])}</div></div>
  <div class="kpi"><div class="kpi-label">AMD</div><div class="kpi-value amd">{dollars(FORECASTS['AMD']['price'])}</div><div class="kpi-sub">Annual model {dollars(PRIOR['AMD'])} · market {dollars(MARKET['AMD'])}</div></div>
  <div class="kpi"><div class="kpi-label">Intel</div><div class="kpi-value intel">{dollars(FORECASTS['Intel']['price'])}</div><div class="kpi-sub">Annual model {dollars(PRIOR['Intel'])} · market {dollars(MARKET['Intel'])}</div></div>
  <div class="kpi"><div class="kpi-label">Intel, kinder margin</div><div class="kpi-value intel">{dollars(sensitivity['price'])}</div><div class="kpi-sub">Starts at the latest {pct(RATIOS['Intel'][-1]['opm'])} margin and glides to 12%</div></div>
</div>
<div class="formula">Value = current book equity + present value of twenty quarterly residual incomes + present value of the terminal value at quarter 20</div>
<div class="formula">Residual income = quarterly net income − (cost of equity / 4) x beginning equity. Terminal value = residual income at quarter 20 x (1 + quarterly growth) / (quarterly cost of equity − quarterly growth).</div>
<h2>Assumptions</h2>
<p>NVIDIA's first quarter is the company's $108bn revenue guide, midpoint, with no China data-center compute. Later quarters decelerate. AMD and Intel step off the latest reported quarter. Operating margin starts at the trailing margin. NVIDIA net income keeps the latest quarter's conversion of operating income. AMD and Intel net income is operating income taxed at 21%.</p>
<div class="cards">
{''.join(assumption_cards)}
</div>
<div class="note">One-time items are left out of the path. NVIDIA's Q1 2026 net margin sits above the operating margin because of equity-security gains, and the cash-flow statement removes $23.7bn of pretax gains in the first half. AMD's listing gain is not treated as ongoing. Intel's quarter includes about a $12.5bn mark on the escrowed-share derivative. Data Center was $89.0bn of the NVIDIA quarter. That is context. It is not a forecast line.</div>
<h2>Five years, summed from the quarters</h2>
<p>Year-1 revenue versus the trailing twelve months is {pct(growth['NVIDIA'])} at NVIDIA, {pct(growth['AMD'])} at AMD, and {pct(growth['Intel'])} at Intel. NVIDIA's year-1 growth is above the old 30% path because the guide is the starting quarter.</p>
{year_table("Year", "rev", "bn").replace("<th></th>", "<th>Revenue</th>", 1)}
{year_table("Year", "ni", "bn").replace("<th></th>", "<th>Net income</th>", 1)}
{year_table("Year", "om", "pct").replace("<th></th>", "<th>Operating margin</th>", 1)}
<h2>Valuation bridge</h2>
<p>Equity each quarter is not a clean-surplus rollforward. Assets equal annualized revenue divided by the latest turnover, and equity is that asset total times the latest equity ratio. Per share uses diluted weighted-average shares: 24,285 million at NVIDIA, 1,659 million at AMD, and 5,104 million at Intel.</p>
{bridge_table(sensitivity)}
<p>Intel's base case glides from a trailing margin near breakeven to 10%, and the value is {dollars(FORECASTS['Intel']['price'])}. Starting at this quarter's {pct(RATIOS['Intel'][-1]['opm'])} operating margin and gliding to 12% produces {dollars(sensitivity['price'])} on equity value of {billions(sensitivity['mve'])}. The prior quarter was an operating loss, so the latest margin is not the run-rate. Either figure is far from the October 5 market price of {dollars(MARKET['Intel'])}.</p>
<h2>Twenty quarters</h2>
<p>Revenue, margin, income, and ending equity are the path. Residual income and its present value use the quarterly cost of equity. The sum of the present-value column is the residual-income term in the bridge.</p>
{''.join(detail)}
<h2>What this page keeps in view</h2>
<p>The DuPont ratios behind the latest quarter and the trailing year are in <a href="dupont-analysis.html">dupont-analysis.html</a>. Quick-pricing coefficients in <a href="quick-pricing-formulas.html">quick-pricing-formulas.html</a> were fit to the annual model and are superseded. The Form 10-K discussion stays in the PDF for strategy. These calculations are the quarters.</p>
<p class="subtitle">Analysis for study purposes, not investment advice.</p>
"""
    return page("Twenty-quarter forecast: NVIDIA, AMD, Intel", body)


def check(sensitivity: dict) -> None:
    expected = {"NVIDIA": 114.08, "AMD": 35.00, "Intel": 2.30}
    for name, price in expected.items():
        if round(FORECASTS[name]["price"], 2) != price:
            raise AssertionError((name, FORECASTS[name]["price"]))
        if len(FORECASTS[name]["quarters"]) != 20:
            raise AssertionError(name)
    if round(sensitivity["price"], 2) != 7.70:
        raise AssertionError(sensitivity["price"])
    if abs(latest_ttm("NVIDIA")["roe"] - 1.172) > 0.002:
        raise AssertionError(latest_ttm("NVIDIA")["roe"])
    if RATIOS["AMD"][-1]["rnoa"] is not None:
        raise AssertionError("AMD modified DuPont should stay open")
    growth = {
        "NVIDIA": 0.544,
        "AMD": 0.270,
        "Intel": 0.189,
    }
    for name, rate in growth.items():
        actual = year_totals(name)[0]["rev"] / latest_ttm(name)["rev"] - 1
        if abs(actual - rate) > 0.001:
            raise AssertionError((name, actual))


def main() -> None:
    sensitivity = forecast("Intel", om0=RATIOS["Intel"][-1]["opm"], om_target=0.12)
    check(sensitivity)
    dupont_path = ROOT / "dupont-analysis.html"
    forecast_path = ROOT / "forecast-valuation.html"
    dupont_path.write_text(dupont_page(), encoding="utf-8")
    forecast_path.write_text(forecast_page(sensitivity), encoding="utf-8")
    print(f"wrote {dupont_path}")
    print(f"wrote {forecast_path}")
    for name in ORDER:
        print(name, f"{FORECASTS[name]['price']:.2f}", billions(FORECASTS[name]["mve"]))
    print("Intel sensitivity", f"{sensitivity['price']:.2f}")


if __name__ == "__main__":
    main()
