#!/usr/bin/env python3
"""NVIDIA vs AMD vs Intel, in the form of the Micron / SK hynix study.

Course-slide DuPont on average balances. The ratio window, the comparison, and
the forecast are quarterly. The 10-K lines still feed the strategy sections.
"""

from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager
import xlsxwriter
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    BaseDocTemplate,
    CondPageBreak,
    Frame,
    HRFlowable,
    Image,
    KeepTogether,
    PageTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
)
from quarterly_model import (
    FORECASTS,
    RATIOS as QRATIOS,
    forecast as quarterly_forecast,
    latest_ttm,
    ttm_at,
    year_totals,
)
from quarterly_statements import QUARTERS, SHARES, SLOTS

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "reports" / "Alam_NVIDIA_Accounting_Performance.pdf"
XLSX = ROOT / "reports" / "NVIDIA_vs_AMD_Intel_DuPont_backup.xlsx"
CHART = ROOT / "reports" / "charts"
CHART.mkdir(parents=True, exist_ok=True)

# Carlito is the metric-compatible face used when Calibri is not installed.
_FONT_SETS = [
    (Path(r"C:\Windows\Fonts"), {
        "Calibri": "calibri.ttf",
        "Calibri-Bold": "calibrib.ttf",
        "Calibri-Italic": "calibrii.ttf",
        "Calibri-BoldItalic": "calibriz.ttf",
    }, "Calibri"),
    (Path("/tmp/fonts/carlito-main/fonts/ttf"), {
        "Calibri": "Carlito-Regular.ttf",
        "Calibri-Bold": "Carlito-Bold.ttf",
        "Calibri-Italic": "Carlito-Italic.ttf",
        "Calibri-BoldItalic": "Carlito-BoldItalic.ttf",
    }, "Carlito"),
]
FONT_FAMILY = "Calibri"
for _dir, _files, _family in _FONT_SETS:
    if all((_dir / name).exists() for name in _files.values()):
        for face, filename in _files.items():
            pdfmetrics.registerFont(TTFont(face, str(_dir / filename)))
            font_manager.fontManager.addfont(str(_dir / filename))
        FONT_FAMILY = _family
        break
else:
    raise SystemExit("No Calibri or Carlito fonts found")

NAVY = colors.HexColor("#1B2A4A")
INK = colors.HexColor("#1C2430")
MUTED = colors.HexColor("#5C6773")
RULE = colors.HexColor("#E3E6EA")
ALT = colors.HexColor("#F6F8FA")
ZEBRA = colors.HexColor("#F3F6F8")
YES = colors.HexColor("#8C2F39")
BASE_BG = colors.HexColor("#EEF1F4")
NV_C = colors.HexColor("#0F6B5C")
AMD_C = colors.HexColor("#D4652F")
INTC_C = colors.HexColor("#1F4E79")
NV_HEX, AMD_HEX, INTC_HEX = "#0F6B5C", "#D4652F", "#1F4E79"
TAX = 0.21

# USD millions. debt = short-term debt + current portion + long-term debt.
# lease = long-term operating lease liability on the balance sheet (Intel: none;
# its operating leases sit in other liabilities, about $0.4bn, and are left out).
# liq = marketable securities (NVIDIA) or short-term investments (AMD, Intel).
# Intel FY2020 trading assets of $15,738 are equity securities, not included in liq.
NV = {
    2021: dict(rev=16675, oi=4532, ni=4332, ix=184, ii=57, ta=28791, eq=16893, cash=847, liq=10714, debt=6963, lease=634, ocf=5822, capex=1128),
    2022: dict(rev=26914, oi=10041, ni=9752, ix=236, ii=29, ta=44187, eq=26612, cash=1990, liq=19218, debt=10946, lease=741, ocf=9108, capex=976),
    2023: dict(rev=26974, oi=4224, ni=4368, ix=262, ii=267, ta=41182, eq=22101, cash=3389, liq=9907, debt=10953, lease=902, ocf=5641, capex=1833),
    2024: dict(rev=60922, oi=32972, ni=29760, ix=257, ii=866, ta=65728, eq=42978, cash=7280, liq=18704, debt=9709, lease=1119, ocf=28090, capex=1069),
    2025: dict(rev=130497, oi=81453, ni=72880, ix=247, ii=1786, ta=111601, eq=79327, cash=8589, liq=34621, debt=8463, lease=1519, ocf=64089, capex=3236),
    2026: dict(rev=215938, oi=130387, ni=120067, ix=259, ii=2300, ta=206803, eq=157293, cash=10605, liq=51951, debt=8468, lease=2572, ocf=102718, capex=6042),
}
AMD = {
    2020: dict(rev=9763, oi=1369, ni=2490, ix=47, ii=8, ta=8962, eq=5837, cash=1595, liq=695, debt=330, lease=201, ocf=1071, capex=294),
    2021: dict(rev=16434, oi=3648, ni=3162, ix=34, ii=8, ta=12419, eq=7497, cash=2535, liq=1073, debt=313, lease=348, ocf=3521, capex=301),
    2022: dict(rev=23601, oi=1264, ni=1320, ix=88, ii=65, ta=67580, eq=54750, cash=4835, liq=1020, debt=2467, lease=396, ocf=3565, capex=450),
    2023: dict(rev=22680, oi=401, ni=854, ix=106, ii=206, ta=67885, eq=55892, cash=3933, liq=1840, debt=2468, lease=535, ocf=1667, capex=546),
    2024: dict(rev=25785, oi=1900, ni=1641, ix=92, ii=182, ta=69226, eq=57568, cash=3787, liq=1345, debt=1721, lease=491, ocf=3041, capex=636),
    # FY25 capex and OCF include discontinued ZT operations ($38 and $1,216).
    2025: dict(rev=34639, oi=3694, ni=4335, ix=131, ii=215, ta=76926, eq=62999, cash=5539, liq=5013, debt=3222, lease=625, ocf=7709, capex=1012),
}
INTC = {
    2020: dict(rev=77867, oi=23678, ni=20899, ix=629, ii=272, ta=153091, eq=81038, cash=5865, liq=2292, debt=36401, lease=0, ocf=35864, capex=14259),
    2021: dict(rev=79024, oi=19456, ni=19868, ix=597, ii=144, ta=168406, eq=95391, cash=4827, liq=24426, debt=38101, lease=0, ocf=29456, capex=18733),
    2022: dict(rev=63054, oi=2334, ni=8017, ix=496, ii=589, ta=182103, eq=103286, cash=11144, liq=17194, debt=42051, lease=0, ocf=15433, capex=24844),
    2023: dict(rev=54228, oi=93, ni=1675, ix=878, ii=1335, ta=191572, eq=109965, cash=7079, liq=17955, debt=49266, lease=0, ocf=11471, capex=25750),
    2024: dict(rev=53101, oi=-11678, ni=-19233, ix=1034, ii=1245, ta=196485, eq=105032, cash=8249, liq=13813, debt=50011, lease=0, ocf=8288, capex=23944),
    2025: dict(rev=52853, oi=-2214, ni=26, ix=1091, ii=1007, ta=211429, eq=126360, cash=14265, liq=23151, debt=46585, lease=0, ocf=9697, capex=14646),
}
BOOKS = {"NVIDIA": NV, "AMD": AMD, "Intel": INTC}
# Ratio years, including the shaded base year at index 0.
WINDOWS = {"NVIDIA": list(range(2021, 2027)), "AMD": list(range(2020, 2026)), "Intel": list(range(2020, 2026))}
ORDER = ("NVIDIA", "AMD", "Intel")
COLORS = {"NVIDIA": NV_C, "AMD": AMD_C, "Intel": INTC_C}
HEXES = {"NVIDIA": NV_HEX, "AMD": AMD_HEX, "Intel": INTC_HEX}


def net_debt(row):
    return row["debt"] + row["lease"] - row["cash"] - row["liq"]


def decompose(book, y0, y1):
    a, b = book[y0], book[y1]
    avg_ta = (a["ta"] + b["ta"]) / 2
    avg_eq = (a["eq"] + b["eq"]) / 2
    nd0, nd1 = net_debt(a), net_debt(b)
    avg_nd = (nd0 + nd1) / 2
    avg_noa = avg_eq + avg_nd
    nfe = (b["ix"] - b["ii"]) * (1 - TAX)
    nopat = b["ni"] + nfe
    npm = b["ni"] / b["rev"]
    ato = b["rev"] / avg_ta
    em = avg_ta / avg_eq
    roe = b["ni"] / avg_eq
    pm = nopat / b["rev"]
    turn = b["rev"] / avg_noa
    rnoa = pm * turn
    near = abs(avg_nd) < 250
    nbc = None if near else nfe / avg_nd
    flev = avg_nd / avg_eq
    if near:
        spread = gain = roe2 = None
    else:
        spread = rnoa - nbc
        gain = spread * flev
        roe2 = rnoa + gain
    alt = b["oi"] * (1 - TAX) / avg_noa
    assert abs(npm * ato * em - roe) < 1e-9
    assert abs(rnoa - nopat / avg_noa) < 1e-9
    if roe2 is not None:
        assert abs(roe2 - roe) < 1e-6
    return dict(
        npm=npm, ato=ato, em=em, roe=roe, pm=pm, turn=turn, rnoa=rnoa, nbc=nbc,
        spread=spread, flev=flev, gain=gain, roe2=roe2, alt=alt, nopat=nopat,
        nfe=nfe, nd=nd1, avg_nd=avg_nd, avg_noa=avg_noa, avg_eq=avg_eq, avg_ta=avg_ta,
        opm=b["oi"] / b["rev"], capex_rev=b["capex"] / b["rev"],
        ocf_capex=b["ocf"] / b["capex"], nd_eq=nd1 / b["eq"],
        ocf_ni=(b["ocf"] / b["ni"] if b["ni"] else None),
    )


RATIOS = {
    name: [decompose(BOOKS[name], y0, y1) for y0, y1 in zip(yrs, yrs[1:])]
    for name, yrs in WINDOWS.items()
}


def pct(x, digits=1):
    if x is None:
        return "n.m."
    return f"{x * 100:.{digits}f}%"


def pct2(x):
    return pct(x, 2)


def xn(x, digits=2):
    if x is None:
        return "n.m."
    return f"{x:.{digits}f}"


def bn(m):
    sign = "-" if m < 0 else ""
    return f"{sign}${abs(m) / 1000:.1f}bn"


def money(m):
    sign = "-" if m < 0 else ""
    return f"{sign}${abs(m):,.0f}"


def latest(name):
    return RATIOS[name][-1]


def row_of(name, year):
    yrs = WINDOWS[name]
    return RATIOS[name][yrs.index(year) - 1]


def qlatest(name):
    return QRATIOS[name][-1]


def qratio(name, slot):
    return QRATIOS[name][SLOTS.index(slot)]


def qline(name, slot):
    """Statement lines for a ratio quarter. Q2 24 is the opening balance and is not a slot."""
    return QUARTERS[name][SLOTS.index(slot) + 1]


def nd_phrase(nd):
    if nd < 0:
        return f"Net cash {bn(abs(nd))}"
    return f"Net debt {bn(nd)}"


# ---------------------------------------------------------------------------
# Excel backup. Ratio cells are formulas off the source block.

def write_excel():
    """Quarterly sources, formula DuPont, trailing comparison, and residual-income forecast."""
    colname = xlsxwriter.utility.xl_col_to_name
    wb = xlsxwriter.Workbook(str(XLSX))
    src = wb.add_worksheet("Sources")
    dup = wb.add_worksheet("DuPont")
    comp = wb.add_worksheet("Comparability")
    cast = wb.add_worksheet("Forecast")
    title = wb.add_format({"bold": True, "font_size": 16, "font_color": "#1B2A4A", "font_name": "Calibri"})
    head = wb.add_format({"bold": True, "font_color": "white", "bg_color": "#1B2A4A", "font_name": "Calibri", "align": "center", "valign": "vcenter", "text_wrap": True})
    head_base = wb.add_format({"bold": True, "font_color": "#1C2430", "bg_color": "#EEF1F4", "font_name": "Calibri", "align": "center", "valign": "vcenter", "text_wrap": True})
    label = wb.add_format({"font_name": "Calibri", "align": "left"})
    label_b = wb.add_format({"font_name": "Calibri", "bold": True})
    num = wb.add_format({"font_name": "Calibri", "num_format": "#,##0;(#,##0)"})
    num_b = wb.add_format({"font_name": "Calibri", "num_format": "#,##0;(#,##0)", "bold": True})
    base = wb.add_format({"font_name": "Calibri", "num_format": "#,##0;(#,##0)", "bg_color": "#EEF1F4"})
    note = wb.add_format({"font_name": "Calibri", "italic": True, "font_color": "#5C6773", "text_wrap": True, "valign": "top"})
    pct_f = wb.add_format({"font_name": "Calibri", "num_format": "0.00%"})
    pct_b = wb.add_format({"font_name": "Calibri", "num_format": "0.00%", "bold": True})
    x_f = wb.add_format({"font_name": "Calibri", "num_format": "0.00"})
    yes_f = wb.add_format({"font_name": "Calibri", "bold": True, "font_color": "#8C2F39", "align": "center"})
    yes_base = wb.add_format({"font_name": "Calibri", "bold": True, "font_color": "#8C2F39", "align": "center", "bg_color": "#EEF1F4"})
    px = wb.add_format({"font_name": "Calibri", "num_format": "0.00", "bold": True})
    co_fmt = {
        "NVIDIA": wb.add_format({"bold": True, "font_color": "white", "bg_color": "#0F6B5C", "align": "center", "font_name": "Calibri"}),
        "AMD": wb.add_format({"bold": True, "font_color": "white", "bg_color": "#D4652F", "align": "center", "font_name": "Calibri"}),
        "Intel": wb.add_format({"bold": True, "font_color": "white", "bg_color": "#1F4E79", "align": "center", "font_name": "Calibri"}),
    }
    labels_q = ["Q2 24 base"] + list(SLOTS)
    lines = [
        ("rev", "Revenue"),
        ("oi", "Operating income"),
        ("ni", "Net income to the company"),
        ("ix", "Interest expense"),
        ("ii", "Interest income"),
        ("ta", "Total assets"),
        ("eq", "Equity, including NCI where presented"),
        ("cash", "Cash"),
        ("liq", "Marketable securities or short-term investments"),
        ("debt", "Debt, including current portion"),
        ("lease", "Long-term operating lease liability"),
        ("ocf", "Operating cash flow"),
        ("capex", "Capital expenditures, cash spent"),
    ]
    src.write(0, 0, "Quarterly source figures, USD millions", title)
    src.write(1, 0, "Shaded columns are the Q2 2024 opening balance and are not ratio quarters. Net debt and NOPAT are formulas. AMD interest income is blank, not zero. The fourth quarter is the rounded fiscal year minus the three rounded quarters.", note)
    src.set_column(0, 0, 62)
    src.set_row(1, 32)
    col_of = {}
    col = 1
    header_row = 3
    src.write(header_row, 0, "Line", head)
    for name in ORDER:
        for i, lab in enumerate(labels_q):
            col_of[(name, i)] = col
            src.write(header_row, col, f"{name} {lab}", head_base if i == 0 else co_fmt[name])
            src.set_column(col, col, 14)
            col += 1
    src.set_row(header_row, 32)
    for r, (key, caption) in enumerate(lines, start=header_row + 1):
        src.write(r, 0, caption, label_b if key in ("rev", "ni", "ta", "eq") else label)
        for name in ORDER:
            for i in range(9):
                c = col_of[(name, i)]
                value = QUARTERS[name][i][key]
                fmt = base if i == 0 else (num_b if key in ("rev", "ni") else num)
                if value is None:
                    src.write_blank(r, c, None, fmt)
                else:
                    src.write_number(r, c, value, fmt)
    r_nd = header_row + 1 + len(lines)
    r_nfe = r_nd + 1
    r_nopat = r_nd + 2
    src.write(r_nd, 0, "Net debt = debt + lease - cash - liquid investments", label_b)
    src.write(r_nfe, 0, "After-tax net interest = (interest expense - interest income) x 0.79", label)
    src.write(r_nopat, 0, "NOPAT = net income + after-tax net interest", label_b)
    key_row = {key: header_row + 1 + i for i, (key, _) in enumerate(lines)}
    for name in ORDER:
        for i in range(9):
            c = col_of[(name, i)]
            cl = colname(c)

            def ref(key, _cl=cl):
                return f"{_cl}{key_row[key] + 1}"

            fmt = base if i == 0 else num
            src.write_formula(r_nd, c, f"={ref('debt')}+{ref('lease')}-{ref('cash')}-{ref('liq')}", fmt)
            src.write_formula(r_nfe, c, f'=IF(COUNTA({ref("ii")})=0,"",({ref("ix")}-{ref("ii")})*(1-0.21))', fmt)
            src.write_formula(r_nopat, c, f'=IF(COUNTA({ref("ii")})=0,"",{ref("ni")}+{cl}{r_nfe + 1})', base if i == 0 else num_b)
    src.write(r_nopat + 2, 0, "Pulled from SEC companyfacts for the Form 10-Q and Form 10-K quarters. A tagged quarter is kept. The fourth quarter equals the rounded fiscal year minus the three rounded quarters, and the plug is within $2 million of the unrounded residual, so four quarters add to the 10-K. AMD interest income is annual-only in the taxonomy, so the cell is blank and COUNTA keeps it from becoming zero. Intel net income is consolidated profit, and Intel equity includes non-controlling interests. Intel cash on 2024-06-29, 2024-09-28, and 2025-03-29 is cash plus restricted cash, because face cash was not tagged. NVIDIA liquid investments from 2026-01-25 are current debt securities plus equity securities at fair value. NVIDIA capex is purchases of property, equipment, and intangibles. AMD and Intel capex are purchases of property and equipment. The annual AMD figure includes $38 million of discontinued capex that is not in these quarters.", note)
    src.set_row(r_nopat + 2, 72)

    dup.write(0, 0, "Quarterly advanced DuPont", title)
    dup.write(1, 0, "Every ratio is a formula on the quarter. YES tests the unrounded quarterly identity. Annualized rows are the quarter times four. Shaded columns are the opening balance. AMD modified DuPont is n.m. because interest income is blank.", note)
    dup.set_row(1, 32)
    dup.set_column(0, 0, 68)
    dup.freeze_panes(4, 1)
    dup.write(3, 0, "RATIOS", head)
    for name in ORDER:
        c0 = col_of[(name, 0)]
        c1 = col_of[(name, 8)]
        dup.merge_range(2, c0, 2, c1, name, co_fmt[name])
        for i, lab in enumerate(labels_q):
            c = col_of[(name, i)]
            dup.write(3, c, lab, head_base if i == 0 else head)
            dup.set_column(c, c, 14)

    def src_ref(name, idx, key):
        return f"Sources!{colname(col_of[(name, idx)])}{key_row[key] + 1}"

    def src_extra(name, idx, row):
        return f"Sources!{colname(col_of[(name, idx)])}{row + 1}"

    ratio_labels = [
        (5, "Traditional DuPont", None),
        (6, "Net margin", "npm"),
        (7, "x Asset turnover (sales / avg total assets)", "ato"),
        (8, "x Leverage (avg assets / avg equity)", "em"),
        (9, "ROE (NI / avg equity)", "roe"),
        (10, "Check: ROE = margin x turnover x leverage", "chk1"),
        (12, "Modified DuPont", None),
        (13, "NOPAT margin", "pm"),
        (14, "x Operating asset turnover (sales / avg NOA)", "turn"),
        (15, "Operating ROA", "rnoa"),
        (16, "Check: operating ROA = NOPAT margin x NOA turn", "chk2"),
        (18, "Operating ROA", "rnoa2"),
        (19, "- Net interest cost (after-tax interest / avg net debt)", "nbc"),
        (20, "Spread", "spread"),
        (21, "x Net financial leverage (avg net debt / avg equity)", "flev"),
        (22, "Gain or loss on financial leverage", "gain"),
        (23, "ROE (NI / avg equity)", "roe2"),
        (24, "Check: ROE = operating ROA + gain", "chk3"),
        (26, "Operating ROA if NOPAT = operating income x 0.79", "alt"),
        (28, "Annualized, quarter x 4. Margins are not annualized.", None),
        (29, "ROE, annualized", "roe_ann"),
        (30, "Asset turnover, annualized", "ato_ann"),
        (31, "Operating ROA, annualized", "rnoa_ann"),
        (32, "Operating ROA on taxed operating income, annualized", "alt_ann"),
    ]
    for r, text, key in ratio_labels:
        style = label_b if key in (None, "roe", "rnoa", "gain", "roe2", "chk1", "chk2", "chk3", "roe_ann", "rnoa_ann") else label
        dup.write(r, 0, text, style)

    for name in ORDER:
        for i in range(9):
            c = col_of[(name, i)]
            if i == 0:
                for r, _text, key in ratio_labels:
                    if key is None:
                        continue
                    dup.write_blank(r, c, None, base)
                continue
            prev = i - 1
            ta0, ta1 = src_ref(name, prev, "ta"), src_ref(name, i, "ta")
            eq0, eq1 = src_ref(name, prev, "eq"), src_ref(name, i, "eq")
            rev = src_ref(name, i, "rev")
            ni = src_ref(name, i, "ni")
            oi = src_ref(name, i, "oi")
            nd0, nd1 = src_extra(name, prev, r_nd), src_extra(name, i, r_nd)
            nopat = src_extra(name, i, r_nopat)
            nfe = src_extra(name, i, r_nfe)
            avg_ta = f"((({ta0})+({ta1}))/2)"
            avg_eq = f"((({eq0})+({eq1}))/2)"
            avg_nd = f"((({nd0})+({nd1}))/2)"
            avg_noa = f"({avg_eq}+{avg_nd})"
            cl = colname(c)
            missing = f'{nopat}=""'
            dup.write_formula(6, c, f"={ni}/{rev}", pct_f)
            dup.write_formula(7, c, f"={rev}/{avg_ta}", x_f)
            dup.write_formula(8, c, f"={avg_ta}/{avg_eq}", x_f)
            dup.write_formula(9, c, f"={ni}/{avg_eq}", pct_b)
            dup.write_formula(10, c, f'=IF(ABS({cl}7*{cl}8*{cl}9-{cl}10)<0.0000005,"YES","NO")', yes_f)
            dup.write_formula(13, c, f'=IF({missing},"n.m.",{nopat}/{rev})', pct_f)
            dup.write_formula(14, c, f'=IF({missing},"n.m.",{rev}/{avg_noa})', x_f)
            dup.write_formula(15, c, f'=IF({missing},"n.m.",{nopat}/{avg_noa})', pct_b)
            dup.write_formula(16, c, f'=IF({missing},"n.m.",IF(ABS({cl}14*{cl}15-{cl}16)<0.0000005,"YES","NO"))', yes_f)
            dup.write_formula(18, c, f'=IF({missing},"n.m.",{cl}16)', pct_f)
            dup.write_formula(19, c, f'=IF(OR({missing},ABS({avg_nd})<250),"n.m.",{nfe}/({avg_nd}))', pct_f)
            dup.write_formula(20, c, f'=IF(OR({missing},{cl}20="n.m."),"n.m.",{cl}19-{cl}20)', pct_f)
            dup.write_formula(21, c, f'=IF({missing},"n.m.",{avg_nd}/{avg_eq})', x_f)
            dup.write_formula(22, c, f'=IF(OR({missing},{cl}21="n.m."),"n.m.",{cl}21*{cl}22)', pct_b)
            dup.write_formula(23, c, f"={ni}/{avg_eq}", pct_b)
            dup.write_formula(24, c, f'=IF(OR({missing},{cl}23="n.m."),"n.m.",IF(ABS({cl}19+{cl}23-{cl}24)<0.0000005,"YES","NO"))', yes_f)
            dup.write_formula(26, c, f"=({oi})*(1-0.21)/({avg_noa})", pct_f)
            dup.write_formula(29, c, f"={cl}10*4", pct_b)
            dup.write_formula(30, c, f"={cl}8*4", x_f)
            dup.write_formula(31, c, f'=IF({cl}16="n.m.","n.m.",{cl}16*4)', pct_b)
            dup.write_formula(32, c, f"={cl}27*4", pct_f)
    dup.write(34, 0, "Averages are wrapped as (((opening)+(closing))/2) so Excel does not read a/b/2 as (a/b)/2. NOPAT uses the statutory 21% rate on net interest only. The taxed-operating-income row is filled in for every company, including AMD. Negative net financial leverage means net cash. Annualized turnover and ROE are the quarterly figures times four. Margins are not annualized. The trailing-twelve-month comparison is on the next sheet.", note)
    dup.set_row(34, 48)

    comp.write(0, 0, "Trailing twelve months, common calendar slot", title)
    comp.write(1, 0, "Flows are the four quarters ending Q2 2026. The opening balance is Q2 2025, the quarter before that window. NVIDIA's period ends about four weeks after AMD and Intel. YES tests net margin x turnover x leverage on those totals.", note)
    comp.set_row(1, 32)
    comp.set_column(0, 0, 62)
    comp.set_column(1, 3, 18)
    comp.write(3, 0, "Line", head)
    for i, name in enumerate(ORDER):
        comp.write(3, 1 + i, name, co_fmt[name])

    def sum_of(name, key, indexes):
        refs = ",".join(src_ref(name, i, key) for i in indexes)
        return f"SUM({refs})"

    def window_formula(name, end_i, kind):
        flows = [end_i - 3, end_i - 2, end_i - 1, end_i]
        open_i = end_i - 4
        rev = sum_of(name, "rev", flows)
        oi = sum_of(name, "oi", flows)
        ni = sum_of(name, "ni", flows)
        ix = sum_of(name, "ix", flows)
        ii = sum_of(name, "ii", flows)
        eq0, eq1 = src_ref(name, open_i, "eq"), src_ref(name, end_i, "eq")
        ta0, ta1 = src_ref(name, open_i, "ta"), src_ref(name, end_i, "ta")
        nd0, nd1 = src_extra(name, open_i, r_nd), src_extra(name, end_i, r_nd)
        avg_eq = f"((({eq0})+({eq1}))/2)"
        avg_ta = f"((({ta0})+({ta1}))/2)"
        avg_nd = f"((({nd0})+({nd1}))/2)"
        avg_noa = f"({avg_eq}+{avg_nd})"
        ii_refs = ",".join(src_ref(name, i, "ii") for i in flows)
        if kind == "rev":
            return f"={rev}"
        if kind == "oi":
            return f"={oi}"
        if kind == "ni":
            return f"={ni}"
        if kind == "opm":
            return f"={oi}/{rev}"
        if kind == "npm":
            return f"={ni}/{rev}"
        if kind == "roe":
            return f"={ni}/{avg_eq}"
        if kind == "ato":
            return f"={rev}/{avg_ta}"
        if kind == "em":
            return f"={avg_ta}/{avg_eq}"
        if kind == "alt":
            return f"={oi}*(1-0.21)/{avg_noa}"
        if kind == "nd":
            return f"={nd1}"
        if kind == "rnoa":
            return f'=IF(COUNTA({ii_refs})<4,"n.m.",({ni}+({ix}-{ii})*(1-0.21))/{avg_noa})'
        raise KeyError(kind)

    latest_rows = [
        (4, "TTM revenue", "rev", num_b),
        (5, "TTM operating income", "oi", num),
        (6, "TTM net income", "ni", num_b),
        (7, "TTM operating margin", "opm", pct_b),
        (8, "TTM net margin", "npm", pct_f),
        (9, "TTM ROE", "roe", pct_b),
        (10, "TTM asset turnover", "ato", x_f),
        (11, "TTM leverage", "em", x_f),
        (13, "TTM operating ROA", "rnoa", pct_b),
        (14, "TTM operating ROA on taxed operating income", "alt", pct_f),
        (15, "Net debt (net cash) at Q2 2026", "nd", num),
    ]
    for r, caption, kind, fmt in latest_rows:
        comp.write(r, 0, caption, label_b if kind in ("rev", "roe", "rnoa") else label)
        for i, name in enumerate(ORDER):
            comp.write_formula(r, 1 + i, window_formula(name, 8, kind), fmt)
    comp.write(12, 0, "Check: TTM ROE = margin x turnover x leverage", label_b)
    for i in range(3):
        cl = colname(1 + i)
        comp.write_formula(12, 1 + i, f'=IF(ABS({cl}9*{cl}11*{cl}12-{cl}10)<0.0000005,"YES","NO")', yes_f)
    comp.write(17, 0, "Trailing ROE by window end", label_b)
    comp.write(18, 0, "Window", head)
    for i, name in enumerate(ORDER):
        comp.write(18, 1 + i, name, co_fmt[name])
    for n, end_i in enumerate(range(4, 9)):
        r = 19 + n
        comp.write(r, 0, SLOTS[end_i - 1], label)
        for i, name in enumerate(ORDER):
            comp.write_formula(r, 1 + i, window_formula(name, end_i, "roe"), pct_f)
    comp.write(25, 0, "Trailing operating margin by window end", label_b)
    for n, end_i in enumerate(range(4, 9)):
        r = 26 + n
        comp.write(r, 0, SLOTS[end_i - 1], label)
        for i, name in enumerate(ORDER):
            comp.write_formula(r, 1 + i, window_formula(name, end_i, "opm"), pct_f)
    comp.write(32, 0, "The latest window's opening equity is the Q2 2025 column on Sources, and the flows are Q3 2025 through Q2 2026. Earlier windows step back one quarter at a time. Operating ROA is n.m. for AMD because interest income is blank.", note)
    comp.set_row(32, 32)

    cast.write(0, 0, "Twenty-quarter residual income", title)
    cast.write(1, 0, "Quarterly revenue, margin, income, and ending equity are values. Residual income, the terminal value, and the price are formulas. Ending equity is not a clean-surplus rollforward: assets are annualized revenue divided by the latest turnover, and equity is that asset total times the latest equity ratio. Cost of equity and terminal growth are the annual model's rates.", note)
    cast.set_row(1, 48)
    cast.set_column(0, 0, 46)
    cast.set_column(1, 8, 16)
    cursor = 3
    bcol = colname(1)
    for name in ORDER:
        fc = FORECASTS[name]
        assumptions = fc["assumptions"]
        ke_r, g_r, gq_r, keq_r, sh_r, bk_r = (cursor + i for i in range(1, 7))
        hdr = cursor + 8
        q0 = cursor + 9
        cast.write(cursor, 0, name, co_fmt[name])
        cast.write(ke_r, 0, "Cost of equity", label)
        cast.write_number(ke_r, 1, assumptions["ke"], pct_f)
        cast.write(g_r, 0, "Terminal growth, annual", label)
        cast.write_number(g_r, 1, assumptions["g"], pct_f)
        cast.write(gq_r, 0, "Terminal growth, quarterly", label)
        cast.write_formula(gq_r, 1, f"=(1+{bcol}{g_r + 1})^(1/4)-1", pct_f)
        cast.write(keq_r, 0, "Cost of equity, quarterly", label)
        cast.write_formula(keq_r, 1, f"={bcol}{ke_r + 1}/4", pct_f)
        cast.write(sh_r, 0, "Diluted weighted-average shares, millions", label)
        cast.write_number(sh_r, 1, fc["shares"], num)
        cast.write(bk_r, 0, "Opening book equity, latest quarter", label_b)
        cast.write_number(bk_r, 1, fc["book"], num_b)
        cast.write(ke_r, 3, "Starting operating margin (trailing)", label)
        cast.write_number(ke_r, 4, fc["om0"], pct_f)
        cast.write(g_r, 3, "Target operating margin", label)
        cast.write_number(g_r, 4, assumptions["om_target"], pct_f)
        cast.write(gq_r, 3, "Net income / operating income", label)
        cast.write_number(gq_r, 4, assumptions["ni_on_oi"], x_f)
        cast.write(keq_r, 3, "Annualized turnover, pinned", label)
        cast.write_number(keq_r, 4, fc["ato"], x_f)
        cast.write(sh_r, 3, "Equity / assets, pinned", label)
        cast.write_number(sh_r, 4, fc["eq_ratio"], pct_f)
        cast.merge_range(bk_r, 3, bk_r, 8, assumptions["note"], note)
        headers = ["Quarter", "Revenue", "Operating margin", "Operating income", "Net income", "Equity, end", "Equity, begin", "Residual income", "Present value"]
        for c, text in enumerate(headers):
            cast.write(hdr, c, text, head)
        for j, q in enumerate(fc["quarters"]):
            r = q0 + j
            excel = r + 1
            cast.write_number(r, 0, q["t"], num)
            cast.write_number(r, 1, q["rev"], num)
            cast.write_number(r, 2, q["om"], pct_f)
            cast.write_number(r, 3, q["oi"], num)
            cast.write_number(r, 4, q["ni"], num)
            cast.write_number(r, 5, q["eq"], num)
            if j == 0:
                cast.write_formula(r, 6, f"={bcol}{bk_r + 1}", num)
            else:
                cast.write_formula(r, 6, f"=F{excel - 1}", num)
            cast.write_formula(r, 7, f"=E{excel}-({bcol}${keq_r + 1}*G{excel})", num)
            cast.write_formula(r, 8, f"=H{excel}/(1+{bcol}${keq_r + 1})^A{excel}", num)
        last = q0 + 19
        pv_r = last + 2
        tv_r = pv_r + 1
        ptv_r = pv_r + 2
        mve_r = pv_r + 3
        px_r = pv_r + 4
        y1_r = pv_r + 5
        cast.write(pv_r, 0, "Present value of residual income", label)
        cast.write_formula(pv_r, 1, f"=SUM(I{q0 + 1}:I{last + 1})", num_b)
        cast.write(tv_r, 0, "Terminal value at quarter 20", label)
        cast.write_formula(tv_r, 1, f"=H{last + 1}*(1+{bcol}{gq_r + 1})/({bcol}{keq_r + 1}-{bcol}{gq_r + 1})", num)
        cast.write(ptv_r, 0, "Present value of terminal value", label)
        cast.write_formula(ptv_r, 1, f"={bcol}{tv_r + 1}/(1+{bcol}{keq_r + 1})^A{last + 1}", num)
        cast.write(mve_r, 0, "Equity value = book + residual income + terminal", label_b)
        cast.write_formula(mve_r, 1, f"={bcol}{bk_r + 1}+{bcol}{pv_r + 1}+{bcol}{ptv_r + 1}", num_b)
        cast.write(px_r, 0, "Value per share", label_b)
        cast.write_formula(px_r, 1, f"={bcol}{mve_r + 1}/{bcol}{sh_r + 1}", px)
        comp_col = colname(1 + ORDER.index(name))
        cast.write(y1_r, 0, "Year-1 revenue / trailing revenue - 1", label)
        cast.write_formula(y1_r, 1, f"=SUM(B{q0 + 1}:B{q0 + 4})/Comparability!{comp_col}5-1", pct_b)
        cursor = y1_r + 3
    sens = quarterly_forecast("Intel", om0=qlatest("Intel")["opm"], om_target=0.12)
    cast.write(cursor, 0, "Intel sensitivity", co_fmt["Intel"])
    cast.write(cursor + 1, 0, "Same twenty quarters and the same 9.7% cost of equity, with the margin starting at the latest quarter and gliding to 12%. This cell is the computed value, not a second formula block.", note)
    cast.set_row(cursor + 1, 32)
    cast.write(cursor + 2, 0, "Sensitivity value per share", label_b)
    cast.write_number(cursor + 2, 1, sens["price"], px)
    cast.write(cursor + 3, 0, "Sensitivity equity value", label)
    cast.write_number(cursor + 3, 1, sens["mve"], num)
    wb.close()


def _style(ax):
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_color("#E3E6EA")
    ax.spines["bottom"].set_color("#E3E6EA")
    ax.tick_params(colors="#3D4654", labelsize=8)
    ax.grid(axis="y", color="#EEF1F4", zorder=0)
    ax.set_axisbelow(True)
    ax.axhline(0, color="#C5CDD6", lw=0.6, zorder=1)


def make_charts():
    plt.rcParams["font.family"] = FONT_FAMILY
    x = list(range(len(SLOTS)))

    def operating_roa(name):
        # AMD does not disclose quarterly interest income, so the comparable
        # operating return is operating income taxed at 21%.
        key = "alt_ann" if name == "AMD" else "rnoa_ann"
        return [row[key] * 100 for row in QRATIOS[name]]

    fig, axes = plt.subplots(1, 3, figsize=(7.35, 2.85), dpi=160)
    panels = [
        ("Operating margin", lambda name: [row["opm"] * 100 for row in QRATIOS[name]]),
        ("ROE, annualized", lambda name: [row["roe_ann"] * 100 for row in QRATIOS[name]]),
        ("Operating ROA, annualized", operating_roa),
    ]
    for ax, (title, series) in zip(axes, panels):
        for name in ORDER:
            ax.plot(x, series(name), color=HEXES[name], marker="o", ms=3.5, lw=1.7, label=name, zorder=3)
        _style(ax)
        ax.set_title(title, fontsize=9, color="#1B2A4A", loc="left", pad=6)
        ax.set_xlim(-0.3, 7.3)
        ax.set_xticks(x)
        ax.set_xticklabels(SLOTS, fontsize=7)
        ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda v, _p: f"{v:.0f}%"))
    handles, labels = axes[0].get_legend_handles_labels()
    fig.legend(handles, labels, loc="upper center", ncol=3, frameon=False, fontsize=8)
    fig.tight_layout(rect=(0, 0, 1, 0.88))
    path = CHART / "performance.png"
    fig.savefig(path)
    plt.close()

    decomp_paths = []
    for name in ORDER:
        roe = [row["roe_ann"] * 100 for row in QRATIOS[name]]
        operating = operating_roa(name)
        fig, ax = plt.subplots(figsize=(2.55, 2.7), dpi=160)
        ax.plot(x, operating, color=HEXES[name], marker="o", ms=3.2, lw=1.6, zorder=3)
        if QRATIOS[name][-1]["gain"] is not None:
            gain = [row["gain"] * 4 * 100 for row in QRATIOS[name]]
            ax.plot(x, gain, color="#8E9BAA", marker="o", ms=3.0, lw=1.3, zorder=3)
        ax.plot(x, roe, color="#1B2A4A", marker="D", ms=3.4, lw=1.2, zorder=4)
        _style(ax)
        ax.set_xticks(x)
        ax.set_xticklabels(SLOTS, fontsize=6.5, rotation=45, ha="right")
        ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda v, _p: f"{v:.0f}%"))
        ax.set_title(name, fontsize=10, color=HEXES[name], loc="left")
        ax.tick_params(axis="y", labelsize=8)
        fig.subplots_adjust(left=0.22, right=0.97, top=0.86, bottom=0.28)
        out = CHART / f"decomp_{name.lower()}.png"
        fig.savefig(out)
        plt.close()
        decomp_paths.append(out)
    return path, decomp_paths


# ---------------------------------------------------------------------------
# PDF

S = {}


def styles():
    common = dict(fontName="Calibri", textColor=INK)
    S["title"] = ParagraphStyle("title", fontName="Calibri-Bold", fontSize=22, leading=26, textColor=NAVY, spaceAfter=1)
    S["sub"] = ParagraphStyle("sub", fontName="Calibri", fontSize=10.5, leading=13.5, textColor=MUTED)
    S["h"] = ParagraphStyle("h", fontName="Calibri-Bold", fontSize=14, leading=17, textColor=NAVY)
    S["deck"] = ParagraphStyle("deck", fontName="Calibri-Italic", fontSize=9, leading=12, textColor=MUTED, spaceBefore=1, spaceAfter=6)
    S["body"] = ParagraphStyle("body", fontName="Calibri", fontSize=9.5, leading=12.6, textColor=INK, spaceAfter=6)
    S["bullet"] = ParagraphStyle("bullet", fontName="Calibri", fontSize=9.5, leading=12.6, textColor=INK, spaceAfter=4, leftIndent=8)
    S["small"] = ParagraphStyle("small", fontName="Calibri", fontSize=8, leading=10.4, textColor=INK)
    S["smallm"] = ParagraphStyle("smallm", fontName="Calibri", fontSize=8, leading=10.4, textColor=MUTED)
    S["note"] = ParagraphStyle("note", fontName="Calibri-Italic", fontSize=8, leading=10.4, textColor=MUTED, spaceBefore=2, spaceAfter=6)
    S["th"] = ParagraphStyle("th", fontName="Calibri-Bold", fontSize=7.5, leading=9.4, textColor=colors.white, alignment=TA_CENTER)
    S["thl"] = ParagraphStyle("thl", fontName="Calibri-Bold", fontSize=7.5, leading=9.4, textColor=colors.white, alignment=TA_LEFT)
    S["td"] = ParagraphStyle("td", fontName="Calibri", fontSize=7.5, leading=9.4, textColor=INK, alignment=TA_LEFT)
    S["tdb"] = ParagraphStyle("tdb", fontName="Calibri-Bold", fontSize=7.5, leading=9.4, textColor=INK, alignment=TA_LEFT)
    S["right"] = ParagraphStyle("right", fontName="Calibri", fontSize=7.5, leading=9.4, textColor=INK, alignment=TA_RIGHT)
    S["rightb"] = ParagraphStyle("rightb", fontName="Calibri-Bold", fontSize=7.5, leading=9.4, textColor=INK, alignment=TA_RIGHT)
    S["cell"] = ParagraphStyle("cell", fontName="Calibri", fontSize=7.4, leading=9.3, textColor=INK, alignment=TA_LEFT)
    S["cellb"] = ParagraphStyle("cellb", fontName="Calibri-Bold", fontSize=7.4, leading=9.3, textColor=INK, alignment=TA_LEFT)
    S["why"] = ParagraphStyle("why", fontName="Calibri", fontSize=7.2, leading=9.1, textColor=INK, alignment=TA_LEFT)
    S["numbox"] = ParagraphStyle("numbox", fontName="Calibri-Bold", fontSize=11, leading=13, textColor=colors.white, alignment=TA_CENTER)
    S["centerw"] = ParagraphStyle("centerw", fontName="Calibri-Bold", fontSize=8, leading=10, textColor=colors.white, alignment=TA_CENTER)
    S["dup"] = ParagraphStyle("dup", fontName="Calibri", fontSize=6.4, leading=7.6, textColor=INK, alignment=TA_RIGHT)
    S["dupb"] = ParagraphStyle("dupb", fontName="Calibri-Bold", fontSize=6.4, leading=7.6, textColor=INK, alignment=TA_RIGHT)
    S["dupl"] = ParagraphStyle("dupl", fontName="Calibri", fontSize=6.6, leading=8.0, textColor=INK, alignment=TA_LEFT)
    S["duplb"] = ParagraphStyle("duplb", fontName="Calibri-Bold", fontSize=6.6, leading=8.0, textColor=INK, alignment=TA_LEFT)
    S["yes"] = ParagraphStyle("yes", fontName="Calibri-Bold", fontSize=8, leading=10, textColor=YES, alignment=TA_CENTER)
    S["dupyes"] = ParagraphStyle("dupyes", fontName="Calibri-Bold", fontSize=6.4, leading=7.6, textColor=YES, alignment=TA_CENTER)
    S["thdark"] = ParagraphStyle("thdark", fontName="Calibri-Bold", fontSize=7.5, leading=9.4, textColor=INK, alignment=TA_CENTER)
    S["banner"] = ParagraphStyle("banner", fontName="Calibri-Bold", fontSize=15, leading=18, textColor=colors.white, alignment=TA_LEFT)
    S["foot"] = ParagraphStyle("foot", **common)
    return S


def P(text, style="body"):
    return Paragraph(text, S[style])


def fitted_image(path, width):
    pixel_w, pixel_h = ImageReader(str(path)).getSize()
    return Image(str(path), width=width, height=width * pixel_h / pixel_w)


def section(number, title, deck):
    num = Paragraph(str(number), S["numbox"])
    head = Paragraph(title, S["h"])
    box = Table([[num, head]], colWidths=[0.32 * inch, 6.9 * inch])
    box.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (0, 0), NAVY),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (0, 0), 0),
        ("RIGHTPADDING", (0, 0), (0, 0), 0),
        ("LEFTPADDING", (1, 0), (1, 0), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("LINEBELOW", (1, 0), (1, 0), 1.25, NAVY),
    ]))
    return KeepTogether([Spacer(1, 8), box, P(deck, "deck")])


def zebra_table(rows, widths, header=True, font=7.5):
    style = [
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LINEBELOW", (0, 0), (-1, -2), 0.3, RULE),
        ("LINEBELOW", (0, -1), (-1, -1), 0.6, colors.HexColor("#C5CDD6")),
    ]
    if header:
        style.append(("BACKGROUND", (0, 0), (-1, 0), NAVY))
        style.append(("TOPPADDING", (0, 0), (-1, 0), 5))
        style.append(("BOTTOMPADDING", (0, 0), (-1, 0), 5))
    start = 1 if header else 0
    for i in range(start, len(rows)):
        if (i - start) % 2 == 1:
            style.append(("BACKGROUND", (0, i), (-1, i), ZEBRA))
    table = Table(rows, colWidths=widths, repeatRows=1 if header else 0)
    table.setStyle(TableStyle(style))
    return table


def snapshot_table():
    columns = []
    for name in ORDER:
        q = qlatest(name)
        ttm = latest_ttm(name)
        ended = qline(name, "Q2 26")["end"]
        columns.append((
            f"Q2 26, ended {ended}",
            bn(qline(name, "Q2 26")["rev"]),
            pct(q["opm"]),
            pct(q["roe_ann"]),
            pct(ttm["roe"]),
            pct(ttm["opm"]),
            nd_phrase(q["nd"]),
        ))
    labels = [
        ("Latest quarter", *(row[0] for row in columns)),
        ("Quarter revenue", *(row[1] for row in columns)),
        ("Quarter operating margin", *(row[2] for row in columns)),
        ("ROE, annualized (quarter x 4)", *(row[3] for row in columns)),
        ("Trailing-twelve-month ROE", *(row[4] for row in columns)),
        ("Trailing-twelve-month operating margin", *(row[5] for row in columns)),
        ("Net cash (net debt), quarter-end", *(row[6] for row in columns)),
    ]
    rows = [[
        P("", "thl"),
        P("NVIDIA", "centerw"),
        P("AMD", "centerw"),
        P("Intel", "centerw"),
    ]]
    for i, (a, b, c, d) in enumerate(labels):
        style = "tdb" if i == 0 else "td"
        rows.append([P(a, style), P(b, "td"), P(c, "td"), P(d, "td")])
    widths = [2.45 * inch, 1.63 * inch, 1.63 * inch, 1.63 * inch]
    table = Table(rows, colWidths=widths)
    table.setStyle(TableStyle([
        ("BACKGROUND", (1, 0), (1, 0), NV_C),
        ("BACKGROUND", (2, 0), (2, 0), AMD_C),
        ("BACKGROUND", (3, 0), (3, 0), INTC_C),
        ("BACKGROUND", (0, 0), (0, 0), NAVY),
        ("BACKGROUND", (0, 1), (-1, 1), colors.white),
        ("BACKGROUND", (0, 2), (-1, 2), ALT),
        ("BACKGROUND", (0, 3), (-1, 3), colors.white),
        ("BACKGROUND", (0, 4), (-1, 4), ALT),
        ("BACKGROUND", (0, 5), (-1, 5), colors.white),
        ("BACKGROUND", (0, 6), (-1, 6), ALT),
        ("BACKGROUND", (0, 7), (-1, 7), colors.white),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LINEBELOW", (0, 1), (-1, -2), 0.25, RULE),
        ("LINEBELOW", (0, -1), (-1, -1), 1, NAVY),
    ]))
    return table


def bullet(lead, rest):
    return P(f"<b>{lead}</b> {rest}", "bullet")


def side_by_side():
    nv, amd, intel = latest("NVIDIA"), latest("AMD"), latest("Intel")
    rows_src = [
        ("What it is",
         "US designer of GPUs, CPUs, networking, and AI systems. Data Center was $193.7bn of FY26 revenue.",
         "US designer of CPUs, GPUs, and adaptive chips. Data Center revenue $16.6bn in FY25; Client and Gaming $14.6bn; Embedded $3.5bn.",
         "US designer and manufacturer. Intel Products operating income $12.7bn in 2025; Intel Foundry operating loss $10.3bn."),
        ("Home and listing", "Santa Clara. Nasdaq: NVDA", "Santa Clara. Nasdaq: AMD", "Santa Clara. Nasdaq: INTC"),
        ("Year-end and rules", "Late January. US GAAP, US dollars. FY26 ended January 25, 2026.", "Last Saturday in December. US GAAP. FY25 ended December 27, 2025.", "Last Saturday in December. US GAAP. FY25 ended December 27, 2025."),
        ("Control", "Widely held.", "Widely held.", "Widely held. Non-controlling interests in Mobileye and the Ireland SCIP partner, inside total equity."),
        ("Auditor", "PricewaterhouseCoopers, PCAOB ID 238, on the FY26 report.", "Ernst & Young, PCAOB ID 42, on the FY25 report.", "Ernst & Young, PCAOB ID 42, on the FY25 report."),
        ("Latest full-year revenue", f"{bn(NV[2026]['rev'])}, 8.0x FY22.", f"{bn(AMD[2025]['rev'])}, up 34% from FY24.", f"{bn(INTC[2025]['rev'])}, down from $79.0bn in FY21."),
        ("Operating margin", pct(nv["opm"]), pct(amd["opm"]), pct(intel["opm"])),
        ("Net income", f"{bn(NV[2026]['ni'])}. Includes about $8.9bn of pretax equity-security gains.", f"{bn(AMD[2025]['ni'])}. Includes a tax benefit and $66m from discontinued operations.", "$26m. Operating loss was $2.2bn. A divestiture gain and a tax charge net out."),
        ("Balance sheet", f"Net cash {bn(abs(nv['nd']))}. Equity {bn(NV[2026]['eq'])}.", f"Net cash {bn(abs(amd['nd']))}. Goodwill $25.1bn and acquisition intangibles $16.7bn.", f"Net debt {bn(intel['nd'])}. Equity {bn(INTC[2025]['eq'])}, including NCI."),
        ("Capital spending", f"{bn(NV[2026]['capex'])}, {pct(nv['capex_rev'], 1)} of revenue. Fabless.", f"{bn(AMD[2025]['capex'])}, including $38m discontinued. Fabless.", f"{bn(INTC[2025]['capex'])} in investing cash flow, {pct(intel['capex_rev'], 0)} of revenue. A further $3.0bn of additions sits in financing."),
        ("AI position", "Sells the AI system. Two direct customers were 22% and 14% of FY26 revenue.", "Sells the merchant alternative. Data Center operating income was $3.6bn. MI308 shipments to China need a US license.", "Product CPUs still earn. The foundry lost $10.3bn. External foundry revenue was $307m."),
        ("Factories", "Wafers from TSMC and Samsung. Memory from SK hynix, Micron, and Samsung. Packaging uses CoWoS.", "TSMC for all 7nm-and-smaller wafers. GlobalFoundries for larger nodes.", "Owns the fabs. Assembly and test include sites in China, among others."),
        ("China", f"Headquarters-based revenue {bn(19677)}, 9% of FY26, after a recast to customer headquarters. Export licenses required for H20 and H200.", f"Geographic revenue {bn(7751)}, 22% of FY25.", f"Geographic billings {bn(12694)}, 24% of FY25. Sales outside the US were 70%."),
        ("Segment disclosure", "Two reportable units, Compute & Networking and Graphics. Segment profit is shown before stock compensation, which is unallocated.", "Data Center, Client and Gaming, and Embedded, each with revenue and operating income.", "Intel Products, Intel Foundry, and an all-other group, with operating income. The foundry loss is visible."),
        ("Guidance in the 10-K", "No numeric quarterly guidance in the annual report.", "No numeric quarterly guidance in the annual report.", "No numeric quarterly guidance in the annual report."),
        ("Extras", "FY26 buybacks $40.4bn. Further $58.5bn authorized at year-end. Groq license and hires: $14.4bn of goodwill.", "FY25 buybacks $1.3bn. $9.4bn still authorized. ZT Systems treated as discontinued.", "Sold 51% of Altera. Partner contributions $5.1bn. Interest capitalized $1.2bn."),
    ]
    header = [P("", "thl"), P("NVIDIA", "centerw"), P("AMD", "centerw"), P("Intel", "centerw")]
    data = [header]
    for label, a, b, c in rows_src:
        data.append([P(label, "cellb"), P(a, "cell"), P(b, "cell"), P(c, "cell")])
    widths = [1.15 * inch, 2.04 * inch, 2.04 * inch, 2.04 * inch]
    table = Table(data, colWidths=widths, repeatRows=1)
    cmds = [
        ("BACKGROUND", (0, 0), (0, 0), NAVY),
        ("BACKGROUND", (1, 0), (1, 0), NV_C),
        ("BACKGROUND", (2, 0), (2, 0), AMD_C),
        ("BACKGROUND", (3, 0), (3, 0), INTC_C),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LINEBELOW", (0, 1), (-1, -2), 0.25, RULE),
        ("LINEBELOW", (0, -1), (-1, -1), 0.8, NAVY),
    ]
    for i in range(1, len(data)):
        if i % 2 == 0:
            cmds.append(("BACKGROUND", (0, i), (-1, i), ZEBRA))
    table.setStyle(TableStyle(cmds))
    return table


def accounting_table():
    topics = [
        ("Revenue",
         "Recognized when control transfers. Two direct customers were 22% and 14% of FY26 revenue. Geographic disclosure switched in Q3 FY26 to the customer's headquarters, and prior years were recast.",
         "Most product revenue is recognized on shipment. Custom products and development services recognized over time were 9% of FY25 revenue, down from 25% in FY23. No customer was 10% of revenue in FY25 or FY24.",
         "Product sales. Three customers were 43% of 2025 revenue. External foundry revenue was $307m, so almost all of the foundry's wafers are an internal transfer.",
         "Concentration can move a quarter at NVIDIA and Intel. AMD's over-time revenue is no longer large enough to dominate the margin."),
        ("Investment gains",
         "Other income was $9.0bn in FY26. The cash-flow statement removes $8.9bn of gains on equity securities. Non-marketable equity securities rose from $3.4bn to $22.3bn after $17.5bn of purchases.",
         "Interest income of $215m and interest expense of $131m. No equivalent equity-security gain in the FY25 net income bridge.",
         "Interest and other was income of $3.3bn. The MD&A describes a $5.6bn pretax gain on the sale of 51% of Altera. The cash-flow statement removes $5.3bn of divestiture gains. A $1.8bn mark-to-market loss on escrowed shares is also in the reconciliation.",
         "Course-slide NOPAT keeps these items, because only interest is added back. NVIDIA's and Intel's net income are not pure operating results in the latest year."),
        ("Inventory",
         "Note 9: inventory provisions of $4.0bn in FY26 cost of revenue and $1.6bn in FY25. The MD&A describes a $4.5bn charge for excess H20 inventory after the April 2025 export rule. Year-end inventory was $21.4bn.",
         "Standard lower-of-cost-or-net-realizable-value accounting. No charge of NVIDIA's or Intel's size in the FY25 discussion.",
         "Inventory reserves have been a recurring charge in the product and accelerator lines. The 2025 MD&A cites lower Gaudi inventory charges as a help to data-center operating income.",
         "Export rules turn inventory into a one-time margin item. Compare operating margin across the license changes, not one year."),
        ("Goodwill and deals",
         "Goodwill rose from $5.2bn to $20.8bn. The Groq license and employee hires were recorded as $14.4bn of goodwill and a $2.5bn technology intangible, cost-to-recreate, five-year life. No customers, products, or equity were bought. Consideration was $13.0bn at closing and $4bn payable within a year, including imputed interest. Goodwill is tax deductible.",
         "Goodwill $25.1bn and acquisition-related intangibles $16.7bn, together 54% of assets. Mostly the Xilinx acquisition. Intangibles still amortize; goodwill does not.",
         "Goodwill $23.9bn, of which Mobileye is $8.2bn. A 2024 test produced a $2.8bn Mobileye impairment. The 2025 test did not. Altera's goodwill left with the divestiture.",
         "NVIDIA's operating-asset base jumped for a workforce and a license that have not yet produced revenue. AMD's turnover is capped by Xilinx. Intel has already shown that a reporting unit can be written down."),
        ("Segments",
         "Compute & Networking and Graphics. Segment operating income of $139.3bn is before $6.4bn of stock compensation and other unallocated costs. Reported operating income is $130.4bn.",
         "Three segments with operating income: Data Center $3.6bn, Client and Gaming $2.9bn, Embedded $1.2bn.",
         "Intel Products operating income $12.7bn against an Intel Foundry operating loss of $10.3bn. The consolidated operating loss of $2.2bn hides a profitable product business.",
         "Intel is the only one of the three where the manufacturing loss is disclosed as its own segment. NVIDIA's segment profit has to be read after the unallocated stock-compensation line."),
        ("Leases and interest",
         "Long-term operating lease liabilities $2.6bn. Leases not yet commenced, expected to start in FY27 through FY30, are a further $22.7bn, mostly data centers. They are not in net debt.",
         "Long-term operating lease liabilities $625m. Small next to equity.",
         "Operating leases are inside other liabilities: $110m current and $281m long-term at the end of 2025. Left out of net debt. Reported interest expense is after $1.2bn capitalized in 2025 and $1.5bn in each of 2024 and 2023.",
         "NVIDIA's lease commitments will become operating assets and liabilities as the data centers start. Intel's capitalized interest keeps reported interest, and therefore net borrowing cost, below the economic cost of the build-out."),
        ("Tax",
         "Tax expense $21.4bn on pretax income of $141.5bn, a 15.1% effective rate. The statutory rate used here on interest is 21%.",
         "Tax benefit of $103m. The rate reconciliation includes an $853m benefit from releasing uncertain tax positions after an IRS determination in April 2025. Effective rate negative.",
         "MD&A: about $9.9bn of non-cash valuation-allowance charges. That is why a $5.6bn gain and a $2.2bn operating loss become $26m of net income.",
         "AMD's FY25 net margin is ahead of after-tax operating profit because of the release. Intel's near-zero net income is a tax and gain result, not a breakeven operation."),
        ("Stock pay",
         "FY26 stock compensation $6.4bn, 4.9% of operating income. Unearned compensation still to be recognized was $14.8bn.",
         "FY25 stock compensation $1.6bn, 44% of operating income of $3.7bn.",
         "FY25 stock compensation $2.4bn, recorded while operating income was a loss.",
         "GAAP operating income already includes it at all three. At AMD it is large relative to profit, so a non-GAAP add-back would change the comparison."),
    ]
    header = [P("Topic", "thl"), P("NVIDIA", "centerw"), P("AMD", "centerw"), P("Intel", "centerw"), P("Why it matters", "centerw")]
    data = [header]
    for topic, a, b, c, why in topics:
        data.append([P(f"<b>{topic}</b>", "cell"), P(a, "cell"), P(b, "cell"), P(c, "cell"), P(why, "why")])
    widths = [0.78 * inch, 1.62 * inch, 1.62 * inch, 1.62 * inch, 1.62 * inch]
    table = Table(data, colWidths=widths, repeatRows=1)
    cmds = [
        ("BACKGROUND", (0, 0), (-1, 0), NAVY),
        ("BACKGROUND", (1, 0), (1, 0), NV_C),
        ("BACKGROUND", (2, 0), (2, 0), AMD_C),
        ("BACKGROUND", (3, 0), (3, 0), INTC_C),
        ("BACKGROUND", (4, 0), (4, 0), NAVY),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LINEBELOW", (0, 1), (-1, -2), 0.25, RULE),
        ("LINEBELOW", (0, -1), (-1, -1), 0.8, NAVY),
    ]
    for i in range(1, len(data)):
        if i % 2 == 0:
            cmds.append(("BACKGROUND", (0, i), (-1, i), ZEBRA))
    table.setStyle(TableStyle(cmds))
    return table


def performance_table():
    """One block per company, eight calendar quarters. Q2 24 is the opening balance and is not shown."""
    blocks = []

    def ocf_text(row):
        if not row["ocf_capex"]:
            return "n.m."
        return f"{row['ocf_capex']:.1f}x"

    metrics = [
        ("Revenue", lambda row, line: bn(line["rev"])),
        ("Operating margin", lambda row, line: pct(row["opm"])),
        ("Net margin", lambda row, line: pct(row["npm"], 1)),
        ("ROE, annualized", lambda row, line: pct(row["roe_ann"])),
        ("Capex / revenue", lambda row, line: pct(row["capex_rev"], 0)),
        ("Operating cash flow / capex", lambda row, line: ocf_text(row)),
        ("Net debt / equity", lambda row, line: pct(row["nd_eq"], 0)),
    ]
    widths = [1.86 * inch] + [0.684 * inch] * 8
    for name in ORDER:
        header = [P(name, "thl")] + [P(slot, "th") for slot in SLOTS]
        data = [header]
        for label, fn in metrics:
            vals = []
            for i, slot in enumerate(SLOTS):
                vals.append(P(fn(QRATIOS[name][i], QUARTERS[name][i + 1]), "rightb" if label == "Revenue" else "right"))
            data.append([P(label, "tdb" if label == "Revenue" else "td")] + vals)
        table = Table(data, colWidths=widths)
        cmds = [
            ("BACKGROUND", (0, 0), (-1, 0), COLORS[name]),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 2),
            ("RIGHTPADDING", (0, 0), (-1, -1), 2),
            ("TOPPADDING", (0, 0), (-1, -1), 2),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
            ("LINEBELOW", (0, 1), (-1, -2), 0.25, RULE),
        ]
        for i in range(1, len(data)):
            if i % 2 == 0:
                cmds.append(("BACKGROUND", (0, i), (-1, i), ZEBRA))
        table.setStyle(TableStyle(cmds))
        blocks.append(KeepTogether([table, Spacer(1, 6)]))
    return blocks


def dupont_blocks():
    """One portrait table per company. The shaded column is the Q2 2024 opening balance."""
    widths = [1.72 * inch] + [0.624 * inch] * 9
    blocks = []
    yes_bg = colors.HexColor("#F7F1F2")
    section_bg = colors.HexColor("#F4F7F8")
    blanks = 9

    def value_row(name, label, key, kind, bold=False):
        row = [P(label, "duplb" if bold else "dupl")]
        row.append(P("", "dup"))
        style = "dupb" if bold else "dup"
        for i in range(len(SLOTS)):
            figure = QRATIOS[name][i][key]
            text = pct2(figure) if kind == "pct" else xn(figure)
            row.append(P(text, style))
        return row

    def yes_row(name, label, modified=False):
        row = [P(label, "dupl"), P("", "dupyes")]
        for i in range(len(SLOTS)):
            ok = (not modified) or QRATIOS[name][i]["rnoa"] is not None
            row.append(P("YES" if ok else "n.m.", "dupyes"))
        return row

    for name in ORDER:
        header = [P(name, "thl"), P("Q2 24", "thdark")]
        header.extend(P(slot, "th") for slot in SLOTS)
        rows = [
            header,
            [P("Traditional DuPont", "duplb")] + [P("", "dup")] * blanks,
            value_row(name, "Net margin", "npm", "pct"),
            value_row(name, "× Asset turnover", "ato", "x"),
            value_row(name, "× Leverage", "em", "x"),
            value_row(name, "ROE", "roe", "pct", True),
            yes_row(name, "Check: margin × turnover × leverage"),
            [P("Modified DuPont", "duplb")] + [P("", "dup")] * blanks,
            value_row(name, "NOPAT margin", "pm", "pct"),
            value_row(name, "× NOA turnover", "turn", "x"),
            value_row(name, "Operating ROA", "rnoa", "pct", True),
            yes_row(name, "Check: NOPAT margin × NOA turn", True),
            value_row(name, "Operating ROA", "rnoa", "pct"),
            value_row(name, "− Net borrowing cost", "nbc", "pct"),
            value_row(name, "Spread", "spread", "pct"),
            value_row(name, "× Net financial leverage", "flev", "x"),
            value_row(name, "Gain or loss on leverage", "gain", "pct", True),
            value_row(name, "ROE", "roe", "pct", True),
            yes_row(name, "Check: operating ROA + gain", True),
            value_row(name, "OI × 0.79 / avg NOA", "alt", "pct"),
            value_row(name, "ROE, annualized × 4", "roe_ann", "pct", True),
            value_row(name, "Operating ROA, annualized × 4", "rnoa_ann", "pct", True),
            value_row(name, "OI × 0.79 ROA, annualized × 4", "alt_ann", "pct"),
        ]
        table = Table(rows, colWidths=widths, repeatRows=1)
        cmds = [
            ("BACKGROUND", (0, 0), (0, 0), COLORS[name]),
            ("BACKGROUND", (2, 0), (-1, 0), COLORS[name]),
            ("BACKGROUND", (0, 1), (0, 1), section_bg),
            ("BACKGROUND", (2, 1), (-1, 1), section_bg),
            ("BACKGROUND", (0, 6), (0, 6), yes_bg),
            ("BACKGROUND", (2, 6), (-1, 6), yes_bg),
            ("BACKGROUND", (0, 7), (0, 7), section_bg),
            ("BACKGROUND", (2, 7), (-1, 7), section_bg),
            ("BACKGROUND", (0, 11), (0, 11), yes_bg),
            ("BACKGROUND", (2, 11), (-1, 11), yes_bg),
            ("BACKGROUND", (0, 18), (0, 18), yes_bg),
            ("BACKGROUND", (2, 18), (-1, 18), yes_bg),
            ("BACKGROUND", (1, 0), (1, -1), BASE_BG),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 2),
            ("RIGHTPADDING", (0, 0), (-1, -1), 2),
            ("TOPPADDING", (0, 0), (-1, -1), 1),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
            ("LINEBELOW", (0, 1), (-1, -2), 0.2, RULE),
            ("LINEABOVE", (0, -1), (-1, -1), 0.6, NAVY),
        ]
        table.setStyle(TableStyle(cmds))
        blocks.append(table)
        blocks.append(Spacer(1, 8))
    return blocks


def fraud_table():
    rows_src = [
        ("Auditor",
         "PwC on the FY26 report. PCAOB ID 238.",
         "Ernst & Young on the FY25 report. PCAOB ID 42.",
         "Ernst & Young on the FY25 report. PCAOB ID 42.",
         "LOW",
         "Top-tier firms on the latest report. I did not re-read every prior opinion page for a change of firm."),
        ("Audit fees",
         "Not in the 10-K. The proxy is incorporated by reference.",
         "Not in the 10-K.",
         "Not in the 10-K.",
         "NO DATA",
         "The proxy statements are not in this filing set."),
        ("Internal controls",
         "Management and PwC concluded disclosure controls and internal control over financial reporting were effective at January 25, 2026. No material weakness is reported. An ERP upgrade is in progress.",
         "The FY25 report includes the auditor's report. I did not find a material-weakness disclosure in the passages read.",
         "Same. No material-weakness sentence in the passages read.",
         "LOW",
         "NVIDIA's conclusion is explicit. AMD and Intel were not re-read line by line for the opinion paragraph."),
        ("Restatements",
         "Nothing in the FY26 statements read as an error correction.",
         "Nothing in the FY25 statements read as an error correction.",
         "Later cash-flow statements revise FY20 operating cash flow from $35,384m as originally filed to $35,864m, and FY21 from $29,991m to $29,456m. I did not find an earnings restatement.",
         "LOW",
         "Use the revised Intel cash-flow comparatives, which this report does. The revision is not an income-statement restatement on the pages read."),
        ("Finance leadership",
         "The FY26 report is signed by the CEO and the CFO. Names and tenure were not tabulated.",
         "Same.",
         "Same.",
         "NO DATA",
         "A tenure check belongs in the proxy."),
        ("Accounting headcount",
         "Not disclosed.",
         "Not disclosed.",
         "Not disclosed.",
         "NO DATA",
         "Not available from the 10-K."),
        ("Reporting complexity",
         "Groq cost-to-recreate valuation. Fair value of $22.3bn of non-marketable equity securities. Export-control inventory. Leases not yet commenced of $22.7bn.",
         "Xilinx intangibles and goodwill. Release of uncertain tax positions. Discontinued ZT Systems. Export licenses for MI308.",
         "Altera gain and valuation allowance. Capitalized interest. SCIP non-controlling interest. Mobileye impairment history. Segment changes.",
         "WATCH",
         "Complexity is highest at Intel and, newly, at NVIDIA. That is a reason to read the notes, not a finding of misstatement."),
        ("Earnings versus cash",
         f"Operating cash flow {bn(NV[2026]['ocf'])} against net income {bn(NV[2026]['ni'])} ({latest('NVIDIA')['ocf_ni']:.2f}x). About $8.9bn of the gap is pretax non-cash investment gains.",
         f"Operating cash flow {bn(AMD[2025]['ocf'])} against net income {bn(AMD[2025]['ni'])} ({latest('AMD')['ocf_ni']:.2f}x), including discontinued operations.",
         f"Operating cash flow {bn(INTC[2025]['ocf'])} against net income of $26m. The ratio is not meaningful. Cash flow does not cover the $14.6bn of investing-section capex.",
         "WATCH",
         "NVIDIA's chip earnings are cash-backed once the gains are removed. Intel's profit figure is not a cash measure of operations."),
        ("Guidance",
         "No numeric quarterly guidance in the 10-K.",
         "No numeric quarterly guidance in the 10-K.",
         "No numeric quarterly guidance in the 10-K.",
         "NO DATA",
         "Outlook lives on the earnings call, which is outside this filing set."),
        ("Disclosure consistency",
         "Geographic revenue was switched to customer headquarters in Q3 FY26 and prior periods were recast. The change is described.",
         "Segment presentation was stable across the years used here.",
         "Segments were reorganized and goodwill was reallocated. The 2024 and 2025 filings describe the changes.",
         "LOW",
         "The changes are disclosed. Intel's segment history is the one that makes a five-year segment comparison hard."),
        ("Related parties and control",
         "Widely held. No controlling shareholder in the statements read.",
         "Small related-party receivable and payable lines in earlier years. Not a feature of FY25.",
         "Ireland SCIP partner contributions were $5.1bn in 2025. Mobileye and other non-controlling interests are part of the equity total. Related-party and partner accounting is a real section of the 10-K.",
         "WATCH",
         "The watch item is Intel's partner structures, not a hidden related-party balance at NVIDIA or AMD."),
        ("Legal and regulatory",
         "Export licenses for China. China's antitrust regulator published a preliminary finding in September 2025 about NVIDIA's compliance with US export controls. H20 and H200 revenue depends on licenses.",
         "MI308 and certain Versal parts need licenses for China. The 10-K says US officials have expressed an expectation of a 15% payment on licensed MI308 sales. That payment is not in effect on the pages read.",
         "Export rules, plus patent cases including VLSI actions in China. CHIPS Act incentives are a large cash item ($1.6bn of capital-related incentives in 2025).",
         "WATCH",
         "The regulatory item that moves the numbers is export control, at all three. It is disclosed."),
    ]
    header = [P("Indicator", "thl"), P("NVIDIA", "centerw"), P("AMD", "centerw"), P("Intel", "centerw"), P("Read", "th")]
    data = [header]
    reads = []
    for ind, a, b, c, read, _why in rows_src:
        data.append([P(f"<b>{ind}</b>", "cell"), P(a, "cell"), P(b, "cell"), P(c, "cell"), P(read, "cellb")])
        reads.append(read)
    widths = [1.05 * inch, 1.85 * inch, 1.55 * inch, 1.85 * inch, 0.7 * inch]
    table = Table(data, colWidths=widths, repeatRows=1)
    read_color = {"LOW": colors.HexColor("#E5F4EE"), "WATCH": colors.HexColor("#FDEFE4"), "NO DATA": colors.HexColor("#EEF1F4")}
    cmds = [
        ("BACKGROUND", (0, 0), (0, 0), NAVY),
        ("BACKGROUND", (1, 0), (1, 0), NV_C),
        ("BACKGROUND", (2, 0), (2, 0), AMD_C),
        ("BACKGROUND", (3, 0), (3, 0), INTC_C),
        ("BACKGROUND", (4, 0), (4, 0), NAVY),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LINEBELOW", (0, 1), (-1, -2), 0.25, RULE),
        ("ALIGN", (4, 1), (4, -1), "CENTER"),
    ]
    for i, read in enumerate(reads, start=1):
        cmds.append(("BACKGROUND", (4, i), (4, i), read_color[read]))
        if i % 2 == 0:
            cmds.append(("BACKGROUND", (0, i), (3, i), ZEBRA))
    table.setStyle(TableStyle(cmds))
    return table


# Half-year and quarter figures from the 2026 Form 10-Qs. Not a sixth DuPont year.
# liq follows the annual definition: NVIDIA marketable debt + marketable equity securities;
# AMD and Intel short-term investments. lease is the long-term operating-lease line.
# Intel interest income and interest expense are the note split, not the interest-and-other total.
Q = {
    "NVIDIA": dict(
        label="Q2 FY27, ended Jul 26, 2026",
        half="Six months ended Jul 26, 2026",
        q_rev=96221, q_oi=63734, q_ni=59688,
        h_rev=177837, h_oi=117270, h_ni=118010,
        h_ix=329, h_ii=1037, h_ocf=74421, h_capex=4434,
        cash=22443, liq=34143 + 42783, debt=1000 + 32366, lease=4985,
        ta=320272, eq=228984,
        filing="Form 10-Q, accession 0001045810-26-000075, filed Aug 26, 2026",
    ),
    "AMD": dict(
        label="Q2 2026, ended Jun 27, 2026",
        half="Six months ended Jun 27, 2026",
        q_rev=11536, q_oi=1990, q_ni=2297,
        h_rev=21789, h_oi=3466, h_ni=3680,
        h_ix=74, h_ii=None, h_ocf=5321, h_capex=1197,
        cash=5086, liq=8025, debt=875 + 2351, lease=1050,
        ta=84464, eq=67224,
        filing="Form 10-Q, accession 0000002488-26-000123, filed Aug 4, 2026",
    ),
    "Intel": dict(
        label="Q2 2026, ended Jun 27, 2026",
        half="Six months ended Jun 27, 2026",
        q_rev=16128, q_oi=1796, q_ni=-10848,
        h_rev=29705, h_oi=-1340, h_ni=-15129,
        h_ix=585, h_ii=667, h_ocf=8102, h_capex=6192,
        cash=12874, liq=16853, debt=1988 + 48549, lease=0,
        ta=202439, eq=103143,
        filing="Form 10-Q, accession 0000050863-26-000157, filed Jul 24, 2026",
    ),
}


def _qnd(name):
    row = Q[name]
    return row["debt"] + row["lease"] - row["cash"] - row["liq"]


def interim_table():
    """Same three-column comparison as the annual snapshot, for the latest quarter and the half."""
    def nd_text(name):
        nd = _qnd(name)
        if nd < 0:
            return f"Net cash {bn(abs(nd))}"
        return f"Net debt {bn(nd)}"

    rows_src = [
        ("Latest quarter on file", Q["NVIDIA"]["label"], Q["AMD"]["label"], Q["Intel"]["label"]),
        ("Quarter revenue", bn(Q["NVIDIA"]["q_rev"]), bn(Q["AMD"]["q_rev"]), bn(Q["Intel"]["q_rev"])),
        ("Quarter operating income", bn(Q["NVIDIA"]["q_oi"]), bn(Q["AMD"]["q_oi"]), bn(Q["Intel"]["q_oi"])),
        ("Quarter operating margin",
         pct(Q["NVIDIA"]["q_oi"] / Q["NVIDIA"]["q_rev"]),
         pct(Q["AMD"]["q_oi"] / Q["AMD"]["q_rev"]),
         pct(Q["Intel"]["q_oi"] / Q["Intel"]["q_rev"])),
        ("Quarter net income", bn(Q["NVIDIA"]["q_ni"]), bn(Q["AMD"]["q_ni"]), bn(Q["Intel"]["q_ni"])),
        ("Half-year revenue", bn(Q["NVIDIA"]["h_rev"]), bn(Q["AMD"]["h_rev"]), bn(Q["Intel"]["h_rev"])),
        ("Half-year operating margin",
         pct(Q["NVIDIA"]["h_oi"] / Q["NVIDIA"]["h_rev"]),
         pct(Q["AMD"]["h_oi"] / Q["AMD"]["h_rev"]),
         pct(Q["Intel"]["h_oi"] / Q["Intel"]["h_rev"])),
        ("Half-year net income", bn(Q["NVIDIA"]["h_ni"]), bn(Q["AMD"]["h_ni"]), bn(Q["Intel"]["h_ni"])),
        ("Net cash (net debt), quarter-end", nd_text("NVIDIA"), nd_text("AMD"), nd_text("Intel")),
    ]
    header = [P("", "thl"), P("NVIDIA", "centerw"), P("AMD", "centerw"), P("Intel", "centerw")]
    data = [header]
    for label, a, b, c in rows_src:
        data.append([P(label, "tdb" if label.startswith("Latest") else "td"), P(a, "td"), P(b, "td"), P(c, "td")])
    widths = [2.15 * inch, 1.70 * inch, 1.70 * inch, 1.70 * inch]
    table = Table(data, colWidths=widths, repeatRows=1)
    cmds = [
        ("BACKGROUND", (0, 0), (0, 0), NAVY),
        ("BACKGROUND", (1, 0), (1, 0), NV_C),
        ("BACKGROUND", (2, 0), (2, 0), AMD_C),
        ("BACKGROUND", (3, 0), (3, 0), INTC_C),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LINEBELOW", (0, 1), (-1, -2), 0.25, RULE),
        ("LINEBELOW", (0, -1), (-1, -1), 1, NAVY),
    ]
    for i in range(1, len(data)):
        if i % 2 == 0:
            cmds.append(("BACKGROUND", (0, i), (-1, i), ZEBRA))
    table.setStyle(TableStyle(cmds))
    return table


def _three(fn):
    return [fn(name) for name in ORDER]


def comparison_table():
    """Trailing comparison on the common calendar slot, plus the latest quarter."""
    header = [P("", "thl"), P("NVIDIA", "centerw"), P("AMD", "centerw"), P("Intel", "centerw")]

    def op_roa(row):
        return row["alt_ann"] if row["rnoa_ann"] is None else row["rnoa_ann"]

    def ttm_op_roa(row):
        return row["alt"] if row["rnoa"] is None else row["rnoa"]

    latest_rows = [
        ("Quarter revenue", lambda name: bn(qline(name, "Q2 26")["rev"])),
        ("Quarter operating margin", lambda name: pct(qlatest(name)["opm"])),
        ("Quarter net margin", lambda name: pct(qlatest(name)["npm"])),
        ("ROE, annualized", lambda name: pct(qlatest(name)["roe_ann"])),
        ("Operating ROA, annualized", lambda name: pct(op_roa(qlatest(name)))),
        ("Trailing revenue", lambda name: bn(latest_ttm(name)["rev"])),
        ("Trailing operating margin", lambda name: pct(latest_ttm(name)["opm"])),
        ("Trailing net margin", lambda name: pct(latest_ttm(name)["npm"])),
        ("Trailing ROE", lambda name: pct(latest_ttm(name)["roe"])),
        ("Trailing operating ROA", lambda name: pct(ttm_op_roa(latest_ttm(name)))),
        ("Net cash (net debt)", lambda name: nd_phrase(qlatest(name)["nd"])),
    ]
    data = [header]
    for label, fn in latest_rows:
        data.append([P(label, "tdb" if "Trailing ROE" == label else "td")] + [P(v, "td") for v in _three(fn)])
    widths = [2.20 * inch, 1.71 * inch, 1.71 * inch, 1.71 * inch]
    table = zebra_table(data, widths)
    # Company colors on the header, over the zebra header navy.
    table.setStyle(TableStyle([
        ("BACKGROUND", (1, 0), (1, 0), NV_C),
        ("BACKGROUND", (2, 0), (2, 0), AMD_C),
        ("BACKGROUND", (3, 0), (3, 0), INTC_C),
    ]))

    trail = [header]
    for slot_index in range(3, len(SLOTS)):
        slot = SLOTS[slot_index]
        trail.append([P(f"Trailing ROE, {slot}", "td")] + [P(pct(ttm_at(name, slot_index)["roe"]), "td") for name in ORDER])
    trail.append([P("Trailing operating margin, Q2 26", "tdb")] + [P(pct(latest_ttm(name)["opm"]), "td") for name in ORDER])
    trail_table = zebra_table(trail, widths)
    trail_table.setStyle(TableStyle([
        ("BACKGROUND", (1, 0), (1, 0), NV_C),
        ("BACKGROUND", (2, 0), (2, 0), AMD_C),
        ("BACKGROUND", (3, 0), (3, 0), INTC_C),
    ]))
    return [table, Spacer(1, 6), trail_table]


def forecast_tables():
    """Five forecast years and the residual-income bridge."""
    market = {"NVIDIA": 236.16, "AMD": 631.57, "Intel": 117.40}
    prior = {"NVIDIA": 54.06, "AMD": 44.38, "Intel": 8.03}
    header = [P("", "thl")] + [P(name, "centerw") for name in ORDER]
    def year_row(label, key, kind):
        vals = []
        for name in ORDER:
            years = year_totals(name)
            if kind == "rev":
                vals.append(bn(years[key]["rev"]))
            elif kind == "ni":
                vals.append(bn(years[key]["ni"]))
            else:
                vals.append(pct(years[key]["om"]))
        return [P(label, "td")] + [P(v, "td") for v in vals]

    rev_rows = [header]
    ni_rows = [header]
    om_rows = [header]
    for k in range(5):
        rev_rows.append(year_row(f"Year {k + 1} revenue", k, "rev"))
        ni_rows.append(year_row(f"Year {k + 1} net income", k, "ni"))
        om_rows.append(year_row(f"Year {k + 1} operating margin", k, "om"))
    growth = [P("Year-1 revenue vs trailing", "tdb")]
    for name in ORDER:
        growth.append(P(pct(year_totals(name)[0]["rev"] / latest_ttm(name)["rev"] - 1), "td"))
    rev_rows.append(growth)
    widths = [2.20 * inch, 1.71 * inch, 1.71 * inch, 1.71 * inch]

    def paint(table):
        table.setStyle(TableStyle([
            ("BACKGROUND", (1, 0), (1, 0), NV_C),
            ("BACKGROUND", (2, 0), (2, 0), AMD_C),
            ("BACKGROUND", (3, 0), (3, 0), INTC_C),
        ]))
        return table

    bridge_rows = [header]
    bridge_src = [
        ("Cost of equity", lambda name: pct(FORECASTS[name]["ke"], 1)),
        ("Terminal growth", lambda name: pct(FORECASTS[name]["g"], 1)),
        ("Opening book equity", lambda name: bn(FORECASTS[name]["book"])),
        ("PV of residual income", lambda name: bn(FORECASTS[name]["pv_ri"])),
        ("PV of terminal value", lambda name: bn(FORECASTS[name]["pv_tv"])),
        ("Equity value", lambda name: bn(FORECASTS[name]["mve"])),
        ("Diluted shares, millions", lambda name: f"{FORECASTS[name]['shares']:,.0f}"),
        ("Value per share", lambda name: f"${FORECASTS[name]['price']:.2f}"),
        ("Prior annual model, per share", lambda name: f"${prior[name]:.2f}"),
        ("Market price, Oct 5, 2026", lambda name: f"${market[name]:.2f}"),
    ]
    for label, fn in bridge_src:
        bold = label in ("Value per share", "Equity value")
        bridge_rows.append([P(label, "tdb" if bold else "td")] + [P(v, "tdb" if bold else "td") for v in _three(fn)])
    tables = [
        paint(zebra_table(rev_rows, widths)),
        paint(zebra_table(ni_rows, widths)),
        paint(zebra_table(om_rows, widths)),
        paint(zebra_table(bridge_rows, widths)),
    ]
    flowables = []
    for table in tables:
        flowables.append(KeepTogether([table]))
        flowables.append(Spacer(1, 6))
    return flowables


def footer(canvas, doc, pagesize):
    canvas.saveState()
    width, height = pagesize
    canvas.setStrokeColor(RULE)
    canvas.line(0.55 * inch, 0.38 * inch, width - 0.55 * inch, 0.38 * inch)
    canvas.setFillColor(MUTED)
    canvas.setFont("Calibri", 8)
    canvas.drawString(0.55 * inch, 0.24 * inch, "NVIDIA vs AMD vs Intel  |  Prepared Oct 6, 2026  |  Analysis for study purposes, not investment advice")
    canvas.drawRightString(width - 0.55 * inch, 0.24 * inch, str(doc.page))
    canvas.restoreState()


def build_story(perf_chart, decomp_charts):
    nv, amd, intel = latest("NVIDIA"), latest("AMD"), latest("Intel")
    nv23 = row_of("NVIDIA", 2023)
    story = []
    story.append(P("NVIDIA vs AMD vs Intel", "title"))
    qn, qa, qi = qlatest("NVIDIA"), qlatest("AMD"), qlatest("Intel")
    tn, ta, ti = latest_ttm("NVIDIA"), latest_ttm("AMD"), latest_ttm("Intel")
    story.append(P("Accounting and performance on quarterly statements, with the 10-K kept for strategy", "sub"))
    story.append(P("Prepared October 6, 2026  ·  Excel backup: NVIDIA_vs_AMD_Intel_DuPont_backup.xlsx", "sub"))
    story.append(P("DuPont, the comparison, and the forecast use eight quarters, Q3 2024 through Q2 2026. Sections 1, 3, 4, 6, and 7 stay on the Form 10-K.", "sub"))
    story.append(Spacer(1, 8))
    story.append(snapshot_table())
    story.append(P("NVIDIA's quarter ends about four weeks after AMD's and Intel's. Annualized ROE is the quarter times four, so it can sit next to an annual figure. Trailing twelve months is the run-rate. Intel's latest net income is the escrowed-share mark. AMD's modified DuPont is open because quarterly interest income is not disclosed.", "note"))
    story.append(Spacer(1, 8))
    story.append(P("Bottom line", "h"))
    story.append(Spacer(1, 3))
    story.append(bullet(
        "Three different businesses, and the quarter already shows it.",
        f"NVIDIA's latest quarter operating margin is {pct(qn['opm'])} on {bn(qline('NVIDIA', 'Q2 26')['rev'])} of revenue. Trailing, it is {pct(tn['opm'])} on {bn(tn['rev'])}. AMD's latest quarter is {pct(qa['opm'])} and the trailing margin is {pct(ta['opm'])}. Intel's latest quarter operating margin is {pct(qi['opm'])}, on operating income of {bn(qline('Intel', 'Q2 26')['oi'])}, while the trailing operating margin is {pct(ti['opm'])}. The 10-K segment split is unchanged: in 2025 the product groups earned $12.7bn and the foundry lost $10.3bn.",
    ))
    story.append(bullet(
        "The quarter-ends are four weeks apart. The return is annualized so it can be read as an annual ROE.",
        f"NVIDIA's quarter ended July 26, 2026. AMD's and Intel's ended June 27, 2026. Annualized ROE this quarter is {pct(qn['roe_ann'])}, {pct(qa['roe_ann'])}, and {pct(qi['roe_ann'])}. Trailing ROE is {pct(tn['roe'])}, {pct(ta['roe'])}, and {pct(ti['roe'])}. Intel's quarterly ROE is the derivative mark, not the operating quarter. The trailing figure is the one to compare.",
    ))
    story.append(bullet(
        "Where the quarter is clean, ROE is still an operating result.",
        f"NVIDIA's net cash of {nd_phrase(qn['nd']).replace('Net cash ', '')} cuts annualized ROE to {pct(qn['roe_ann'])} from an annualized operating ROA of {pct(qn['rnoa_ann'])}. AMD's modified DuPont is not computed. Traditional ROE is. Intel's latest ROE is not the operating result.",
    ))
    story.append(bullet(
        "Do not forecast the one-time lines.",
        f"NVIDIA's Q1 2026 net margin was {pct(qratio('NVIDIA', 'Q1 26')['npm'])} against an operating margin of {pct(qratio('NVIDIA', 'Q1 26')['opm'])}, which is the equity-security gain. AMD's other income lifts net income above operating income, and interest income is not split out. Intel's latest quarter net loss of {bn(abs(qline('Intel', 'Q2 26')['ni']))} sits on an operating profit. The forecast in section 8 does not repeat those items.",
    ))
    story.append(bullet(
        "Fraud screen: nothing found that points to fraud.",
        "The items worth diligence are NVIDIA's new goodwill and investment gains, Intel's non-operating profit and capitalized interest, and AMD's tax release. Details are in the addendum. That screen is still the 10-K.",
    ))
    story.append(Spacer(1, 4))
    story.append(P("How to read the periods", "h"))
    story.append(Spacer(1, 2))
    story.append(P("The ratio window is eight quarters, Q3 2024 through Q2 2026, paired on the calendar. NVIDIA's period ends about four weeks after AMD's and Intel's. The shaded Q2 2024 column is the opening balance and is not a ratio quarter. Sections 1, 3, 4, 6, and 7 stay on the Form 10-K. They are the strategic record. In those sections, NVIDIA FY26, ended January 25, 2026, is the neighbor of AMD and Intel FY25, ended December 27, 2025.", "body"))
    story.append(P("The identity is the quarter. Net margin times asset turnover times leverage equals ROE, on that quarter's income and the average of the opening and closing balance. Annualized figures are the quarter times four. A quarterly ROE near 28% is an annualized ROE near 112%. Trailing twelve months sums four quarters and divides by the average of equity at the quarter before that window and equity at the end. Margins are not annualized.", "body"))
    story.append(P("Modified DuPont adds back after-tax net interest at the 21% statutory rate. AMD does not disclose quarterly interest income, so NOPAT, net borrowing cost, and the leverage gain are n.m. The row that taxes operating income at 21% is filled in for all three. Intel's net income is consolidated profit, not the attributable line. Intel's equity includes non-controlling interests.", "body"))
    story.append(Spacer(1, 2))
    story.append(P("Contents", "h"))
    contents = [
        ("1", "Preview: where the three sit in the AI stack"),
        ("2", "NVIDIA, AMD, and Intel side by side"),
        ("3", "Strategy"),
        ("4", "Accounting"),
        ("5", "Performance, DuPont and modified DuPont"),
        ("6", "Miscellaneous: three things worth knowing about each"),
        ("7", "Addendum: fraud red-flag screen"),
        ("8", "Twenty-quarter residual-income forecast"),
        ("", "Sources and method"),
    ]
    crow = []
    for num, title in contents:
        crow.append([P(num, "tdb"), P(title, "td")])
    ctable = Table(crow, colWidths=[0.4 * inch, 6.8 * inch])
    ctable.setStyle(TableStyle([
        ("LINEBELOW", (0, 0), (-1, -2), 0.4, RULE),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING", (0, 0), (0, -1), 2),
    ]))
    story.append(ctable)

    story.append(CondPageBreak(1.7 * inch))
    story.append(section("1", "Preview: where the three sit in the AI stack",
                         "NVIDIA names its own place in the stack. AMD and Intel occupy the neighboring boxes."))
    story.append(P("Four layers, from the FY26 NVIDIA 10-K and the peer filings. This report's ratios cover the two designers and the one manufacturer. The memory suppliers are named by NVIDIA and are not analyzed here.", "body"))
    stack = [[
        P("Layer", "thl"), P("Who", "thl"), P("What the filing says", "thl"), P("Who buys it", "thl"),
    ], [
        P("<b>1 Design</b>", "cell"), P("NVIDIA, AMD", "cellb"),
        P("GPUs, CPUs, networking, and full systems. NVIDIA also sells the software stack around the chip.", "cell"),
        P("Cloud providers, OEMs, and AI labs, usually through direct customers.", "cell"),
    ], [
        P("<b>2 Manufacture</b>", "cell"), P("TSMC, Samsung, Intel", "cellb"),
        P("NVIDIA buys wafers from TSMC and Samsung. AMD buys 7nm-and-smaller wafers from TSMC, and larger nodes from GlobalFoundries. Intel runs its own fabs.", "cell"),
        P("The designers, and, for Intel Foundry, a thin external book.", "cell"),
    ], [
        P("<b>3 Memory</b>", "cell"), P("SK hynix, Micron, Samsung", "cellb"),
        P("NVIDIA says it purchases memory from these three. Not in the ratio tables.", "cell"),
        P("NVIDIA, and the server makers.", "cell"),
    ], [
        P("<b>4 Buyer</b>", "cell"), P("A few large customers", "cellb"),
        P("NVIDIA: two direct customers were 22% and 14% of FY26 revenue. Intel: three customers were 43%. AMD: no 10% customer in FY25.", "cell"),
        P("Their own clients. Those clients are not given statements here.", "cell"),
    ]]
    stable = Table(stack, colWidths=[0.85 * inch, 1.45 * inch, 3.15 * inch, 1.8 * inch])
    stable.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), NAVY),
        ("BACKGROUND", (0, 1), (-1, 1), colors.HexColor("#E7F3EF")),
        ("BACKGROUND", (0, 3), (-1, 3), ALT),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LINEBELOW", (0, 1), (-1, -2), 0.25, RULE),
    ]))
    story.append(stable)
    story.append(Spacer(1, 6))
    story.append(P("How the three companies touch each other", "h"))
    story.append(bullet("NVIDIA and AMD draw on the same foundry.", "AMD's 10-K says a failure by TSMC to manufacture its 7nm-and-smaller wafers would have a material effect. NVIDIA names TSMC and Samsung the same way. They compete for capacity as well as for customers."))
    story.append(bullet("Intel is trying to be the factory and still be the product.", "In 2025 the product groups made $12.7bn of operating income and the foundry lost $10.3bn. External foundry revenue was $307m. The foundry's customer is still mostly Intel."))
    story.append(bullet("Memory is a supplier, not a peer in these ratios.", "NVIDIA lists SK hynix, Micron, and Samsung as memory vendors and says it uses CoWoS packaging. That is the link to the memory report. It is not a number in the DuPont."))
    story.append(bullet("Profit in the latest 10-K sits with the designer that does not own the fab.", f"Operating margin {pct(nv['opm'])} at NVIDIA, {pct(amd['opm'])} at AMD, and {pct(intel['opm'])} at Intel. The quarterly comparison is in the next section."))

    story.append(section("2", "NVIDIA, AMD, and Intel side by side",
                         "The first table is the latest Form 10-K, the strategic snapshot. The quarterly comparison follows."))
    story.append(side_by_side())
    story.append(Spacer(1, 8))
    story.append(P("Quarterly comparison", "h"))
    story.append(P("Same calendar slot. NVIDIA's quarter ends about four weeks later. Where AMD's operating ROA is shown, it is operating income taxed at 21%, because quarterly interest income is not disclosed.", "deck"))
    story.extend(comparison_table())
    story.append(P("Trailing ROE uses four quarters of net income over the average of equity at the quarter before that window and equity at the window end. The first trailing row is Q2 2025.", "note"))
    story.append(P("High-level differences", "h"))
    story.append(bullet("NVIDIA sells a system and is paid like a software company.", f"Trailing operating margin is {pct(tn['opm'])} on {bn(tn['rev'])}. The latest quarter is {pct(qn['opm'])} on {bn(qline('NVIDIA', 'Q2 26')['rev'])}, and annualized ROE is {pct(qn['roe_ann'])}. Capex is still {pct(qn['capex_rev'])} of the quarter. The balance sheet is net cash."))
    story.append(bullet("AMD sells the alternative and still carries Xilinx.", f"Trailing operating margin is {pct(ta['opm'])} and trailing ROE is {pct(ta['roe'])}. Annualized asset turnover in the latest quarter is {xn(qa['ato_ann'])}x. Goodwill and acquisition intangibles were 54% of assets in the FY25 10-K, which is why turnover cannot look like NVIDIA's."))
    story.append(bullet("Intel's latest quarter is not the turnaround.", f"Trailing operating margin is {pct(ti['opm'])} and trailing ROE is {pct(ti['roe'])}. The latest quarter's {pct(qi['opm'])} operating margin sits on a net loss. The 10-K still shows capex at {pct(intel['capex_rev'], 0)} of FY25 revenue, and net income of $26m was not evidence that the turnaround had arrived."))

    story.append(section("3", "Strategy",
                         "The strategic question is different for each firm. NVIDIA's is how long a 60% margin lasts. AMD's is whether the GPU franchise can earn on the Xilinx asset base. Intel's is whether the foundry ever earns its capex."))
    story.append(P("Shared economics", "h"))
    story.append(bullet("The customer is building AI factories.", "All three sell into that build-out. NVIDIA sells the accelerator and the system around it. AMD sells a second source and the x86 CPU. Intel sells the CPU and is spending to sell wafers."))
    story.append(bullet("Leading-edge supply is concentrated.", "NVIDIA and AMD both name TSMC. A year of tight wafers raises their revenue and their inventory risk at the same time. Intel's bet is that owning the fab is the way out of that queue."))
    story.append(bullet("Export control is now part of the product plan.", "NVIDIA's H20 and H200, and AMD's MI308, ship to China only with a license. Intel's China billings are still a quarter of revenue. A rule change shows up as an inventory charge, not only as a missed order."))
    story.append(Spacer(1, 2))
    headers = [
        (NV_C, "NVIDIA: keep the platform, return the cash"),
        (AMD_C, "AMD: be the second source, digest Xilinx"),
        (INTC_C, "Intel: fund the fab from the product business"),
    ]
    texts = [
        [
            "The product is the rack and the software. Data Center was $193.7bn of $215.9bn. Two direct customers were 22% and 14% of revenue.",
            "Manufacturing stays at TSMC and Samsung. NVIDIA keeps the design, the networking, and the software installed on the chips.",
            "Cash came back as buybacks, $40.4bn in FY26, with $58.5bn still authorized. Alongside that, $17.5bn went into non-marketable equity securities.",
            "The Groq license and the employees hired with it added $14.4bn of goodwill and no product revenue yet.",
        ],
        [
            "Data Center revenue grew 32% to $16.6bn and operating income was $3.6bn. Client and Gaming grew 51% to $14.6bn. Embedded, at $3.5bn, was roughly flat.",
            "The 10-K says TSMC produces every 7nm-and-smaller microprocessor and GPU wafer. That dependence is stated as a risk, not a footnote.",
            "Buybacks were $1.3bn, with $9.4bn still authorized. The balance-sheet job is earning a return on $25.1bn of goodwill and $16.7bn of acquisition intangibles.",
            "MI308 sales into China depend on licenses. Geographic revenue from China was $7.8bn, 22% of the year.",
        ],
        [
            "Intel Products made $12.7bn of operating income. Intel Foundry lost $10.3bn. External foundry revenue was $307m, so the product groups are still the foundry's customer.",
            "Investing-section capex was $14.6bn against $9.7bn of operating cash flow. Capital incentives returned $1.6bn. SCIP partner contributions were $5.1bn.",
            "The sale of 51% of Altera produced cash and a gain that does not recur. The next year's test is operating income.",
            "Three customers are 43% of revenue. That is a narrow set of buyers from which to fill new fabs. Reported interest is after $1.2bn of capitalization.",
        ],
    ]
    col_tables = []
    white = ParagraphStyle("wh", fontName="Calibri-Bold", fontSize=8, leading=10.2, textColor=colors.white)
    body_s = S["cell"]
    for (color, title), paras in zip(headers, texts):
        data = [[Paragraph(title, white)]] + [[P(p, "cell")] for p in paras]
        t = Table(data, colWidths=[2.38 * inch])
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), color),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 5),
            ("RIGHTPADDING", (0, 0), (-1, -1), 5),
            ("TOPPADDING", (0, 0), (-1, 0), 5),
            ("BOTTOMPADDING", (0, 0), (-1, 0), 5),
            ("TOPPADDING", (0, 1), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 1), (-1, -1), 3),
            ("BACKGROUND", (0, 1), (-1, -1), colors.HexColor("#F8F9FB")),
        ]))
        col_tables.append(t)
    wrap = Table([col_tables], colWidths=[2.42 * inch] * 3)
    wrap.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 1),
        ("RIGHTPADDING", (0, 0), (-1, -1), 1),
    ]))
    story.append(wrap)
    story.append(Spacer(1, 6))
    story.append(P("What to watch", "h"))
    story.append(bullet("Does the 60% operating margin survive a looser supply of accelerators?", "FY23 is the reminder inside this window: NVIDIA's operating margin was 15.7% and ROE was 17.9%. The investment gains will not fill that gap if the product margin falls."))
    story.append(bullet("What does $14.4bn of Groq goodwill earn?", "There is no revenue in the purchase accounting. The test is whether the next 10-K can point at a product. The intangible amortizes over five years. The goodwill does not."))
    story.append(bullet("Can AMD's data-center profit grow faster than the Xilinx asset base sits still?", "Turnover of 0.47x is the acquisition. Margin recovery without turnover recovery leaves ROE in single digits, which is where it is."))
    story.append(bullet("Does Intel Foundry's loss shrink before the capex has to be earned?", "Operating cash flow does not cover investing-section capex. Capitalized interest flatters the interest line while the fabs are built. The Altera gain will not be in next year's net income."))

    story.append(section("4", "Accounting",
                         "All three report under US GAAP. The differences that move the ratios are in the right-hand column."))
    story.append(accounting_table())
    story.append(Spacer(1, 6))
    story.append(P("Quality of earnings", "h"))
    story.append(bullet(
        "NVIDIA: operations are cash-backed, and the latest net income is not only operations.",
        f"Operating cash flow was {bn(NV[2026]['ocf'])} against net income of {bn(NV[2026]['ni'])}. The cash-flow statement removes $8.9bn of pretax gains on equity securities. Take those out and the cash conversion of the rest is close. Receivables and inventory grew with Blackwell. The provisions are disclosed.",
    ))
    story.append(bullet(
        "AMD: cash flow is stronger than net income, and net income is helped by a tax release.",
        f"Operating cash flow was {bn(AMD[2025]['ocf'])}, of which $1.2bn was discontinued operations, against net income of {bn(AMD[2025]['ni'])}. The $853m tax benefit is not cash. Operating margin of {pct(amd['opm'])} is the cleaner number for the year.",
    ))
    story.append(bullet(
        "Intel: do not read $26m of net income as breakeven.",
        f"Operating loss {bn(abs(INTC[2025]['oi']))}. The Altera gain and the valuation allowance dominate the path from that loss to $26m. Operating cash flow of {bn(INTC[2025]['ocf'])} is real and is still short of the $14.6bn investing-section capex. Course-slide operating ROA in 2025 is {pct(intel['rnoa'], 2)} because NOPAT keeps the gain and the tax charge. Taxing operating income at 21% instead, on the same net operating assets, operating ROA is {pct(intel['alt'])}.",
    ))
    story.append(P("The quarterly ratios use the reported figures. The Excel row that taxes operating income at 21% sits beside course-slide operating ROA. AMD's course-slide row is blank. Section 5 discusses the gap. The 10-K quality points above are unchanged.", "note"))

    story.append(section("5", "Performance, DuPont and modified DuPont",
                         "Eight quarters on one calendar. Returns in the chart are annualized. The identity in the table is the quarter."))
    story.append(fitted_image(perf_chart, 7.25 * inch))
    story.append(P("The slots are Q3 2024 through Q2 2026. Operating margin is the quarter. ROE and operating ROA are the quarter times four. AMD's operating ROA in the third panel is operating income taxed at 21%. NVIDIA's quarter ends about four weeks later than the slot label shared with AMD and Intel.", "note"))
    story.extend(performance_table())
    story.append(P("Capex is cash spent, sign ignored. A single quarter is lumpy, so the coverage comment below uses the trailing four quarters. AMD's quarterly capex is purchases of property and equipment. The $38m of discontinued capex is in the FY25 10-K total and is not in these quarters. Intel's quarterly capex is the investing section. The further financing-section equipment additions stay in the 10-K discussion and are not in this ratio. Net debt is debt plus the long-term operating lease line, minus cash and liquid investments. Negative means net cash.", "note"))
    story.append(bullet(
        "The last four quarters did not converge.",
        f"Against the year-ago quarter, revenue is up {pct(qline('NVIDIA', 'Q2 26')['rev'] / qline('NVIDIA', 'Q2 25')['rev'] - 1)} at NVIDIA, {pct(qline('AMD', 'Q2 26')['rev'] / qline('AMD', 'Q2 25')['rev'] - 1)} at AMD, and {pct(qline('Intel', 'Q2 26')['rev'] / qline('Intel', 'Q2 25')['rev'] - 1)} at Intel. Trailing revenue is {bn(tn['rev'])}, {bn(ta['rev'])}, and {bn(ti['rev'])}.",
    ))
    story.append(bullet(
        "NVIDIA's margin dipped once and then widened. AMD recovered from an operating loss. Intel's trailing margin is about zero.",
        f"NVIDIA's operating margin was {pct(qratio('NVIDIA', 'Q1 25')['opm'])} in Q1 2025, the export-control quarter, and {pct(qn['opm'])} in the latest quarter. Trailing, it is {pct(tn['opm'])}. AMD posted an operating loss of {bn(abs(qline('AMD', 'Q2 25')['oi']))} in Q2 2025 and an operating margin of {pct(qa['opm'])} in the latest quarter. Trailing margin is {pct(ta['opm'])}. Intel's latest operating margin is {pct(qi['opm'])}. Trailing, it is {pct(ti['opm'])}.",
    ))
    story.append(bullet(
        "Capital intensity still says who owns the factory.",
        f"Over the last four quarters, NVIDIA spent {bn(tn['capex'])} and operating cash flow covered it {tn['ocf'] / tn['capex']:.1f} times. AMD spent {bn(ta['capex'])} and covered it {ta['ocf'] / ta['capex']:.1f} times. Intel spent {bn(ti['capex'])} against operating cash flow of {bn(ti['ocf'])}, {ti['ocf'] / ti['capex']:.2f} times coverage.",
    ))
    story.append(bullet(
        "Balance sheets, at the latest quarter-end.",
        f"NVIDIA holds net cash of {bn(abs(qn['nd']))}. AMD holds net cash of {bn(abs(qa['nd']))}. Intel carries net debt of {bn(qi['nd'])}, {pct(qi['nd_eq'], 0)} of quarter-end equity. Debt at NVIDIA rose with the June 2026 notes. The 10-K still has Intel's FY24 net-debt peak and the Altera proceeds. Those are history, not this quarter.",
    ))

    story.append(Spacer(1, 6))
    story.append(P("Advanced DuPont decomposition of return on equity", "h"))
    story.append(P("One table per company. The shaded column is Q2 2024, used only for the averages. YES is the unrounded quarterly identity. Annualized rows are the quarter times four and are not a second identity.", "deck"))
    story.extend(dupont_blocks())
    story.append(P("Equity is the average of beginning and ending balance-sheet equity. Intel's total includes non-controlling interests. Net income is consolidated profit at Intel. NOPAT adds back only after-tax net interest at 21%. AMD's interest income is not disclosed by quarter, so those rows are n.m. The operating-income row is filled in for all three. Net debt is debt plus the long-term operating lease line, minus cash and liquid investments. Negative net financial leverage means net cash. YES is the unrounded quarter. Printed figures are rounded to two decimals. Current operating leases sit inside accrued liabilities and are not in net debt. Intel's operating leases, about $0.4bn in the 10-K, are in other liabilities and are not in net debt. NVIDIA's non-marketable equity securities stay inside net operating assets.", "note"))

    story.append(P("Reading the decomposition", "h"))
    story.append(Spacer(1, 2))
    imgs = [fitted_image(p, 2.38 * inch) for p in decomp_charts]
    img_row = Table([imgs], colWidths=[2.42 * inch] * 3)
    img_row.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("LEFTPADDING", (0, 0), (-1, -1), 2),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2),
    ]))
    story.append(img_row)
    story.append(P("Color is annualized operating ROA. At AMD it is operating income taxed at 21%, and there is no gain line. Gray, where it is drawn, is the annualized leverage gain. Navy diamonds are annualized ROE. The scales differ.", "note"))
    story.append(bullet(
        "Where the quarter is clean, ROE is an operating result. Intel's latest quarter is not clean.",
        f"NVIDIA's leverage gain this quarter is {pct(qn['gain'])} of ROE. Annualized, operating ROA is {pct(qn['rnoa_ann'])} and ROE is {pct(qn['roe_ann'])}. The gap is the cash drag. AMD's modified DuPont is open. Annualized ROE is {pct(qa['roe_ann'])}, and operating income taxed at 21% earns {pct(qa['alt_ann'])} annualized. Intel's latest leverage gain is {pct(qi['gain'])}, on a quarter whose net income is the escrowed-share mark.",
    ))
    story.append(bullet(
        "At NVIDIA, the margin held and the operating-asset base got heavier.",
        f"Latest-quarter net margin is {pct(qn['npm'])}, next to an operating margin of {pct(qn['opm'])}. Annualized operating ROA was {pct(qratio('NVIDIA', 'Q2 25')['rnoa_ann'])} a year earlier and is {pct(qn['rnoa_ann'])} now. Ending net operating assets, equity plus net debt, are {bn(qline('NVIDIA', 'Q2 26')['eq'] + qn['nd'])}. The 10-K is still the source for what sits inside the operating assets: $22.3bn of non-marketable equity securities at FY26, $21.4bn of inventory, and the Groq goodwill. The product margin did not fall.",
    ))
    story.append(bullet(
        "AMD's ROE broke when Xilinx arrived. The goodwill is still the reason turnover is low.",
        f"On the 10-K history, FY21 ROE was {pct(row_of('AMD', 2021)['roe'])} and FY22 ROE was {pct(row_of('AMD', 2022)['roe'])} after assets jumped from {bn(AMD[2021]['ta'])} to {bn(AMD[2022]['ta'])}. This quarter, annualized asset turnover is {xn(qa['ato_ann'])}x and annualized ROE is {pct(qa['roe_ann'])}. Trailing, taxing operating income at 21% and dividing by average net operating assets gives {pct(ta['alt'])}. Course-slide operating ROA is not computed, because interest income is not split out of other income.",
    ))
    story.append(bullet(
        "Intel's trailing operating ROA is a loss. The latest quarter's operating ROA is the derivative, not the foundry.",
        f"Trailing operating ROA is {pct(ti['rnoa'])} and trailing ROE is {pct(ti['roe'])}. The latest quarter, annualized, is an operating ROA of {pct(qi['rnoa_ann'])} because net income is the mark. Operating margin that quarter is {pct(qi['opm'])}. Taxing that operating income at 21% gives an annualized return of {pct(qi['alt_ann'])}. The FY25 10-K operating ROA of {pct(intel['rnoa'], 2)} kept the Altera gain. It is not the trailing figure.",
    ))
    story.append(bullet(
        "Reported borrowing cost still understates Intel's build-out.",
        f"This quarter, after-tax net interest over average net debt is {pct(qi['nbc'])}. That uses reported interest. The FY25 10-K interest expense of $1,091m is after $1.2bn capitalized into the fabs, and the quarterly line is the same accounting. Net financial leverage this quarter is {xn(qi['flev'])}.",
    ))
    story.append(bullet(
        "Idle cash is still NVIDIA's ROE question.",
        f"The quarterly yield in the modified-DuPont borrowing cost is {pct(qn['nbc'])}, against a quarterly operating ROA of {pct(qn['rnoa'])}. Annualized operating ROA is {pct(qn['rnoa_ann'])}. The 10-K uses of cash, the $40.4bn of buybacks and the $17.5bn of equity-security purchases, plus the leases not yet started, are what will move the drag. Borrowing will not.",
    ))

    misc_head = section("6", "Miscellaneous", "Three things about each company that the statements make hard to miss.")
    misc_headers = [
        (NV_C, "NVIDIA"),
        (AMD_C, "AMD"),
        (INTC_C, "Intel"),
    ]
    misc_items = [
        [
            ("<b>1. It paid $14.4bn of goodwill for a license and a team</b>",
             "The Groq note is explicit: no customer contracts, no existing products, no equity interest. Goodwill of $14.4bn, a $2.5bn technology intangible with a five-year life, $13.0bn paid at closing, and $4bn payable within a year including imputed interest. The goodwill is tax deductible. Pro forma revenue was not presented because the effect was not material."),
            ("<b>2. It bought $17.5bn of stakes and marked a gain of $8.9bn</b>",
             "Non-marketable equity securities ended at $22.3bn, up from $3.4bn. Unrealized gains in the rollforward were $2.4bn. The cash-flow statement removes $8.9bn of gains on non-marketable and publicly held equity securities from operating cash flow. Other income of $9.0bn is the income-statement home of that amount."),
            ("<b>3. $22.7bn of data-center leases are not on the balance sheet yet</b>",
             "They are expected to commence between FY27 and FY30. Long-term operating lease liabilities today are $2.6bn. Net debt will look different when those leases start, even if no debt is issued. Two direct customers, at 22% and 14% of revenue, are the demand those leases assume."),
        ],
        [
            ("<b>1. More than half the assets are the Xilinx deal</b>",
             "Goodwill $25.1bn and acquisition-related intangibles $16.7bn, out of assets of $76.9bn. The intangibles amortize. The goodwill does not, until a test fails. Asset turnover of 0.47x is that fact, not a statement about the factories AMD does not own."),
            ("<b>2. An IRS letter added $853m to net income</b>",
             "FY25 tax was a benefit of $103m. The reconciliation identifies an $853m benefit from releasing uncertain tax positions after reasonable-cause relief on dual consolidated losses, approved in April 2025. Operating income of $3.7bn is the figure that does not depend on that letter."),
            ("<b>3. It spent the year both shipping and waiting on a license</b>",
             "MI308 products for China required a license. Some licenses were granted, and shipments started in the fourth quarter. The 10-K also records an expressed official expectation that the US government will receive 15% of licensed MI308 revenue. It says that arrangement is not in effect. China was $7.8bn of geographic revenue."),
        ],
        [
            ("<b>1. Net income of $26m, operating loss of $2.2bn</b>",
             "The MD&A puts a $5.6bn pretax gain on the sale of 51% of Altera inside interest and other. The cash-flow statement removes $5.3bn of divestiture gains. About $9.9bn of non-cash valuation allowances went through tax. Consolidated net income is $26m. Net income attributable to Intel is a loss of $267m, because non-controlling interests were allocated income of $293m."),
            ("<b>2. The interest line is not the cost of the fabs</b>",
             "Interest expense of $1,091m is net of $1.2bn capitalized in 2025, and of $1.5bn in each of the two prior years. Cash paid for interest, net of capitalized interest, was $1,106m. As long as the fabs are under construction, reported net borrowing cost will look too low."),
            ("<b>3. The partner is now a financing source</b>",
             "Partner contributions were $5.1bn in 2025 and $12.7bn in 2024. A further $3.0bn of property and equipment additions is classified in financing, not in investing. Reading only the $14.6bn investing line understates the cash going into equipment. Non-controlling interests on the balance sheet are the equity side of these structures."),
        ],
    ]
    misc_cols = []
    for (color, title), items in zip(misc_headers, misc_items):
        data = [[Paragraph(title, white)]]
        for head, text in items:
            data.append([P(head, "cellb")])
            data.append([P(text, "cell")])
        t = Table(data, colWidths=[2.38 * inch])
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), color),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 5),
            ("RIGHTPADDING", (0, 0), (-1, -1), 5),
            ("TOPPADDING", (0, 0), (-1, 0), 4),
            ("BOTTOMPADDING", (0, 0), (-1, 0), 4),
            ("TOPPADDING", (0, 1), (-1, -1), 2),
            ("BOTTOMPADDING", (0, 1), (-1, -1), 2),
            ("BACKGROUND", (0, 1), (-1, -1), colors.HexColor("#FBFCFD")),
        ]))
        misc_cols.append(t)
    mwrap = Table([misc_cols], colWidths=[2.42 * inch] * 3)
    mwrap.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 1),
        ("RIGHTPADDING", (0, 0), (-1, -1), 1),
    ]))
    story.append(KeepTogether([misc_head, mwrap]))

    story.append(section("7", "Addendum: fraud red-flag screen",
                         "A screen of the annual reports, not an audit. It can raise a question. It cannot rule fraud in or out."))
    story.append(fraud_table())
    story.append(Spacer(1, 6))
    story.append(P("Conclusion", "h"))
    story.append(bullet("No indicator here points to fraud at any of the three.", "The latest opinions are from PwC and Ernst & Young. NVIDIA's management and auditor report effective internal control at the FY26 year-end. Cash flow supports NVIDIA's and AMD's earnings once one-time gains and the AMD tax benefit are identified."))
    story.append(bullet("Intel warrants more diligence than the other two, for structural reasons.", "Net income is a gain and a tax charge. Interest is capitalized. A partner funds part of the build-out and shares in equity. Segments have been reorganized. None of that is hidden. All of it sits between operating income and the net income a DuPont will pick up."))
    story.append(bullet("NVIDIA's new item is the Groq goodwill and the equity-security book.", "Both are described in the notes, with methods and amounts. The open question is economic, not forensic: what those assets earn."))
    story.append(bullet("Not answered by this screen.", "Audit fees, accounting headcount, and a page-by-page check that the auditor did not change during the five years. The first two are outside the 10-K. The third can be done from the opinion pages and was not."))

    sens = quarterly_forecast("Intel", om0=qi["opm"], om_target=0.12)
    story.append(CondPageBreak(3.2 * inch))
    story.append(section(
        "8",
        "Twenty-quarter residual-income forecast",
        "Five years, one quarter at a time, off the latest quarter. Cost of equity and terminal growth are the annual model's rates.",
    ))
    story.append(P("Residual income each quarter is net income minus the quarterly cost of equity times beginning equity. The horizon is 20 quarters, so the terminal value is not doing the work of a two-year window. Cost of equity is 13.5% for NVIDIA, 12.5% for AMD, and 9.7% for Intel. Terminal growth is 3% a year, converted to a quarterly rate. Those are the rates already used on the annual model. This section does not adopt the higher-growth, lower-discount scenarios in the separate buy-and-sell note.", "body"))
    story.append(P("The path starts at the latest quarter, not at the last 10-K. NVIDIA's next quarter is the company's $108bn revenue guide, midpoint, with no China data-center compute. Later quarters decelerate from there. AMD and Intel step off the latest quarter. Operating margin starts at the trailing margin, which is the run-rate, and glides to 55% at NVIDIA, 22% at AMD, and 10% at Intel. NVIDIA net income is operating income times the latest quarter's net-income-to-operating-income ratio. AMD and Intel net income is operating income taxed at 21%.", "body"))
    story.append(bullet(
        "One-time items are not repeated.",
        "NVIDIA's Q1 2026 net margin is above the operating margin because of equity-security gains. The cash-flow statement removes $23.7bn of pretax gains in the first half. AMD's other income, including a listing gain, is not treated as ongoing, and quarterly interest income is not disclosed. Intel's quarter includes the mark on the escrowed-share derivative, about $12.5bn in the quarter, which is not interest and is not forecast. Data Center was $89.0bn of the NVIDIA quarter. That is context. It is not a line in the forecast.",
    ))
    story.append(bullet(
        "Equity is not a clean-surplus rollforward.",
        f"Assets each quarter equal annualized revenue divided by the latest quarter's turnover, {xn(FORECASTS['NVIDIA']['ato'])}x at NVIDIA, {xn(FORECASTS['AMD']['ato'])}x at AMD, and {xn(FORECASTS['Intel']['ato'])}x at Intel. Equity is that asset total times the latest equity ratio. Dividends are not modeled. The simplification is the same one the annual model used.",
    ))
    story.append(P("Per share uses diluted weighted-average shares for the latest quarter: 24,285 million at NVIDIA, 1,659 million at AMD, and 5,104 million at Intel. Intel was in a loss, so diluted shares are close to basic.", "body"))
    story.extend(forecast_tables())
    story.append(Spacer(1, 6))
    story.append(bullet(
        "The value moved because the base moved.",
        f"The quarterly model puts NVIDIA at ${FORECASTS['NVIDIA']['price']:.2f}, AMD at ${FORECASTS['AMD']['price']:.2f}, and Intel at ${FORECASTS['Intel']['price']:.2f}. The prior annual model, on the last 10-K and the same costs of equity, was $54.06, $44.38, and $8.03. Year-1 revenue versus the trailing twelve months is {pct(year_totals('NVIDIA')[0]['rev'] / tn['rev'] - 1)} at NVIDIA, {pct(year_totals('AMD')[0]['rev'] / ta['rev'] - 1)} at AMD, and {pct(year_totals('Intel')[0]['rev'] / ti['rev'] - 1)} at Intel. NVIDIA's year-1 growth is above the old 30% path because the $108bn guide is the starting quarter, not a linear step off FY26 revenue of {bn(NV[2026]['rev'])}.",
    ))
    story.append(bullet(
        "Intel stays below book on these assumptions, including a kinder margin path.",
        f"Trailing margin is about breakeven, so the base case glides from there to 10% and the value is ${FORECASTS['Intel']['price']:.2f}. If the path instead starts at this quarter's {pct(qi['opm'])} operating margin and glides only to 12%, the value is ${sens['price']:.2f}. That quarter's margin is not the run-rate: the prior quarter was an operating loss. Either figure is far from the October 5 market price of $117.40.",
    ))
    story.append(bullet(
        "Market prices are a snapshot, not a new download.",
        "On October 5, 2026 the buy-and-sell note used $236.16 for NVIDIA, $631.57 for AMD, and $117.40 for Intel. This model does not update those prices and does not change that note's verdicts. The gap between these values and those prices is the point of publishing the quarterly base. It is not a recommendation.",
    ))

    story.append(Spacer(1, 8))
    story.append(P("Sources and method", "h"))
    story.append(Spacer(1, 3))
    story.append(P("Quarterly figures", "h"))
    story.append(P("The ratio tables, the comparison, the forecast, and the Excel identities use quarterly statement lines, in USD millions. A tagged quarter is kept. A year-to-date total is split by subtracting the prior year-to-date total. The fourth quarter is the rounded fiscal year minus the three rounded quarters, so the four quarters add to the 10-K. That plug is within $2 million of the unrounded residual. The Excel DuPont sheet computes every ratio, and the YES cells test the quarterly identity. Averages are written as (((opening)+(closing))/2).", "body"))
    story.append(P("NVIDIA flows are revenue, operating income, net income, interest expense, interest income, operating cash flow, and purchases of property, equipment, and intangibles. Liquid investments are marketable securities through October 2025, and current debt securities plus equity securities at fair value from January 25, 2026. AMD interest income is annual-only in the filing taxonomy, so it is blank and modified DuPont is not computed. AMD capex is purchases of property and equipment. The FY25 10-K capex of $1,012m includes $38m discontinued, which is not in the quarterly series. Intel net income is consolidated profit, the line that ties to FY25 income of $26m, not net income attributable to Intel. Intel equity includes non-controlling interests. Intel capex is investing-section additions. Face cash was not tagged on June 29, 2024, September 28, 2024, and March 29, 2025, so those three balances are cash plus restricted cash.", "body"))
    story.append(P("The latest filings in the ratio window are NVIDIA's Form 10-Q for the quarter ended July 26, 2026, filed August 26, 2026, and the AMD and Intel Form 10-Qs for the quarter ended June 27, 2026, filed August 4 and July 24. Shares are diluted weighted-average shares for that quarter.", "body"))
    story.append(P("10-K figures, used for strategy", "h"))
    story.append(P("Sections 1, 3, 4, 6, and 7, and the first table in section 2, use the Form 10-K lines. They are not the ratio window. NVIDIA FY22 through FY26, AMD FY20 through FY25, and Intel FY20 through FY25 are the statement lines named in the prior source note: income, cash, liquid investments, debt, leases, equity, assets, operating cash flow, and capex. Intel consolidated net income is the total, not the attributable line. Intel FY20 and FY21 operating cash flow is the revised comparative. AMD FY25 capex of $1,012m and operating cash flow of $7,709m include discontinued operations.", "body"))
    story.append(P("Context in the strategy sections, each from the latest 10-K: NVIDIA's Groq note, segments, customers, geography, inventory, lease commitments, and buybacks. AMD's segments, TSMC dependence, China revenue, the tax reconciliation, and goodwill. Intel's segment operating income, the Altera gain, capitalized interest, partner contributions, and the financing-section equipment additions. The half-year one-time amounts in section 8, the equity-security gains, the escrowed-share mark, and Data Center revenue of $89.0bn, are note disclosures in the latest 10-Qs. They are not DuPont inputs.", "body"))
    story.append(P("Limits", "h"))
    story.append(P("Quarter-ends are about four weeks apart. Pairing by calendar slot is the comparison. It is not a same-day comparison. In the strategy tables, NVIDIA FY26 is still paired with AMD and Intel FY25, because those are the latest 10-Ks.", "body"))
    story.append(P("Net debt leaves out current operating lease liabilities and Intel's operating leases of about $0.4bn. It leaves NVIDIA's non-marketable equity securities inside net operating assets. AMD's modified DuPont is open on purpose. The forecast pins turnover and the equity ratio to the latest quarter and does not roll equity forward from earnings. Moving those choices would change the level. The ordering of the three companies would not.", "body"))
    story.append(P("Audit fees, headcount, and a full auditor-tenure check were not done. Market prices are the October 5, 2026 snapshot in the buy-and-sell note, not a live quote. The buy-and-sell verdicts are unchanged and use different assumptions.", "body"))
    story.append(P("This is analysis for study purposes. It is not investment advice, and I am not a financial advisor.", "body"))
    return story


def main():
    styles()
    write_excel()
    perf, decomps = make_charts()
    OUT.parent.mkdir(parents=True, exist_ok=True)
    doc = BaseDocTemplate(
        str(OUT),
        title="NVIDIA vs AMD vs Intel",
        author="Accounting and performance study",
    )
    frame_p = Frame(0.58 * inch, 0.52 * inch, letter[0] - 1.16 * inch, letter[1] - 0.78 * inch, id="p", showBoundary=0)
    doc.addPageTemplates([
        PageTemplate(id="portrait", frames=[frame_p], onPage=lambda c, d: footer(c, d, letter), pagesize=letter),
    ])
    doc.build(build_story(perf, decomps))
    nv = latest("NVIDIA")
    q = qlatest("NVIDIA")
    assert abs(nv["roe"] - (nv["rnoa"] + nv["gain"])) < 1e-9
    assert abs((q["rnoa"] + q["gain"]) - q["roe"]) < 1e-6
    assert qlatest("AMD")["rnoa"] is None
    assert abs(latest_ttm("NVIDIA")["roe"] - 1.172) < 0.002
    assert len(FORECASTS["NVIDIA"]["quarters"]) == 20
    print("wrote", OUT)
    print("wrote", XLSX)
    print(
        "quarterly price",
        f"{FORECASTS['NVIDIA']['price']:.2f}",
        f"{FORECASTS['AMD']['price']:.2f}",
        f"{FORECASTS['Intel']['price']:.2f}",
    )
    print("NVIDIA TTM ROE", pct(latest_ttm("NVIDIA")["roe"]), "annualized latest", pct(q["roe_ann"]))


if __name__ == "__main__":
    main()
