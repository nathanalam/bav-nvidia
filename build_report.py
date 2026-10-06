#!/usr/bin/env python3
"""NVIDIA vs AMD vs Intel, in the form of the Micron / SK hynix study.

Course-slide DuPont on average balances. Figures are the 10-K statement lines
tied in the source notes at the bottom of the PDF and in the Excel backup.
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

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "reports" / "Alam_NVIDIA_Accounting_Performance.pdf"
XLSX = ROOT / "reports" / "NVIDIA_vs_AMD_Intel_DuPont_backup.xlsx"
CHART = ROOT / "reports" / "charts"
CHART.mkdir(parents=True, exist_ok=True)

FONT = Path(r"C:\Windows\Fonts")
pdfmetrics.registerFont(TTFont("Calibri", str(FONT / "calibri.ttf")))
pdfmetrics.registerFont(TTFont("Calibri-Bold", str(FONT / "calibrib.ttf")))
pdfmetrics.registerFont(TTFont("Calibri-Italic", str(FONT / "calibrii.ttf")))
pdfmetrics.registerFont(TTFont("Calibri-BoldItalic", str(FONT / "calibriz.ttf")))
font_manager.fontManager.addfont(str(FONT / "calibri.ttf"))
font_manager.fontManager.addfont(str(FONT / "calibrib.ttf"))

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


# ---------------------------------------------------------------------------
# Excel backup. Ratio cells are formulas off the source block.

def write_excel():
    wb = xlsxwriter.Workbook(str(XLSX))
    src = wb.add_worksheet("Sources")
    dup = wb.add_worksheet("DuPont")
    title = wb.add_format({"bold": True, "font_size": 16, "font_color": "#1B2A4A", "font_name": "Calibri"})
    head = wb.add_format({"bold": True, "font_color": "white", "bg_color": "#1B2A4A", "font_name": "Calibri", "align": "center", "valign": "vcenter", "text_wrap": True})
    label = wb.add_format({"font_name": "Calibri", "align": "left"})
    label_b = wb.add_format({"font_name": "Calibri", "bold": True})
    num = wb.add_format({"font_name": "Calibri", "num_format": "#,##0;(#,##0)"})
    num_b = wb.add_format({"font_name": "Calibri", "num_format": "#,##0;(#,##0)", "bold": True})
    base = wb.add_format({"font_name": "Calibri", "num_format": "#,##0;(#,##0)", "bg_color": "#EEF1F4"})
    note = wb.add_format({"font_name": "Calibri", "italic": True, "font_color": "#5C6773", "text_wrap": True})
    pct_f = wb.add_format({"font_name": "Calibri", "num_format": "0.00%"})
    pct_b = wb.add_format({"font_name": "Calibri", "num_format": "0.00%", "bold": True})
    pct_base = wb.add_format({"font_name": "Calibri", "num_format": "0.00%", "bg_color": "#EEF1F4"})
    x_f = wb.add_format({"font_name": "Calibri", "num_format": "0.00"})
    x_base = wb.add_format({"font_name": "Calibri", "num_format": "0.00", "bg_color": "#EEF1F4"})
    yes_f = wb.add_format({"font_name": "Calibri", "bold": True, "font_color": "#8C2F39", "align": "center"})
    yes_base = wb.add_format({"font_name": "Calibri", "bold": True, "font_color": "#8C2F39", "align": "center", "bg_color": "#EEF1F4"})
    co_fmt = {
        "NVIDIA": wb.add_format({"bold": True, "font_color": "white", "bg_color": "#0F6B5C", "align": "center", "font_name": "Calibri"}),
        "AMD": wb.add_format({"bold": True, "font_color": "white", "bg_color": "#D4652F", "align": "center", "font_name": "Calibri"}),
        "Intel": wb.add_format({"bold": True, "font_color": "white", "bg_color": "#1F4E79", "align": "center", "font_name": "Calibri"}),
    }

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
    src.write(0, 0, "Source figures, USD millions", title)
    src.write(1, 0, "Tied to the 10-K statement lines named in the PDF sources. Base years are shaded. Net debt and the ratios are formulas.", note)
    src.set_column(0, 0, 52)
    src.set_row(1, 32)
    # Map (company, year) -> column on the source sheet. Row of each line is fixed.
    col_of = {}
    col = 1
    header_row = 3
    src.write(header_row, 0, "Line", head)
    for name in ORDER:
        for i, year in enumerate(WINDOWS[name]):
            col_of[(name, year)] = col
            fmt = base if i == 0 else num
            src.write(header_row, col, f"{name} FY{year % 100:02d}", co_fmt[name] if i else head)
            src.set_column(col, col, 16, fmt)
            col += 1
    for r, (key, caption) in enumerate(lines, start=header_row + 1):
        src.write(r, 0, caption, label_b if key in ("rev", "ni", "ta", "eq") else label)
        for name in ORDER:
            for i, year in enumerate(WINDOWS[name]):
                c = col_of[(name, year)]
                value = BOOKS[name][year][key]
                src.write_number(r, c, value, base if i == 0 else (num_b if key in ("rev", "ni") else num))
    # Formula rows
    r_nd = header_row + 1 + len(lines)
    r_nfe = r_nd + 1
    r_nopat = r_nd + 2
    src.write(r_nd, 0, "Net debt = debt + lease - cash - liquid investments", label_b)
    src.write(r_nfe, 0, "After-tax net interest = (interest expense - interest income) x 0.79", label)
    src.write(r_nopat, 0, "NOPAT = net income + after-tax net interest", label_b)
    key_row = {key: header_row + 1 + i for i, (key, _) in enumerate(lines)}
    for name in ORDER:
        for i, year in enumerate(WINDOWS[name]):
            c = col_of[(name, year)]
            cl = xlsxwriter.utility.xl_col_to_name(c)

            def ref(key, _cl=cl):
                return f"{_cl}{key_row[key] + 1}"

            shade = base if i == 0 else num
            src.write_formula(r_nd, c, f"={ref('debt')}+{ref('lease')}-{ref('cash')}-{ref('liq')}", shade)
            src.write_formula(r_nfe, c, f"=({ref('ix')}-{ref('ii')})*(1-0.21)", shade)
            src.write_formula(r_nopat, c, f"={ref('ni')}+{cl}{r_nfe + 1}", num_b if i else base)

    src.write(r_nopat + 2, 0, "Intel FY2020 and FY2021 operating cash flow is the revised comparative in the later cash-flow statement (originally $35,384 and $29,991). Intel FY2020 trading assets of $15,738 are not in liquid investments. AMD FY2025 capex of $1,012 and operating cash flow of $7,709 include discontinued operations ($38 and $1,216). Intel capex is the investing-section additions, not the further additions classified in financing ($3,026 in 2025 and $1,178 in 2024). Current operating lease liabilities are inside accrued liabilities and are not in net debt.", note)
    src.set_row(r_nopat + 2, 48)

    # DuPont sheet
    dup.write(0, 0, "NVIDIA vs AMD vs Intel: advanced DuPont", title)
    dup.write(1, 0, "Every ratio is a formula. YES tests the unrounded identity. Shaded columns are the base year and are not ratio years. n.m. is not used; average net debt is never near zero in this panel.", note)
    dup.set_row(1, 32)
    dup.set_column(0, 0, 62)
    dup.freeze_panes(4, 1)

    # Column layout mirrors the source columns so formulas can point across.
    dup.write(3, 0, "RATIOS", head)
    for name in ORDER:
        years = WINDOWS[name]
        c0 = col_of[(name, years[0])]
        c1 = col_of[(name, years[-1])]
        dup.merge_range(2, c0, 2, c1, name, co_fmt[name])
        for year in years:
            c = col_of[(name, year)]
            dup.write(3, c, f"FY{year % 100:02d}", head)
            dup.set_column(c, c, 12)

    def src_ref(name, year, key):
        c = xlsxwriter.utility.xl_col_to_name(col_of[(name, year)])
        return f"Sources!{c}{key_row[key] + 1}"

    def src_extra(name, year, row):
        c = xlsxwriter.utility.xl_col_to_name(col_of[(name, year)])
        return f"Sources!{c}{row + 1}"

    # Ratio rows. Base column stays blank (shaded) except we still shade it.
    # For a ratio year, beginning balances are the prior column on this sheet's source.
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
        (19, "− Net interest cost (after-tax interest / avg net debt)", "nbc"),
        (20, "Spread", "spread"),
        (21, "x Net financial leverage (avg net debt / avg equity)", "flev"),
        (22, "Gain or loss on financial leverage", "gain"),
        (23, "ROE (NI / avg equity)", "roe2"),
        (24, "Check: ROE = operating ROA + gain", "chk3"),
        (26, "Operating ROA if NOPAT = operating income x 0.79", "alt"),
    ]
    for r, text, _key in ratio_labels:
        style = label_b if _key in (None, "roe", "rnoa", "gain", "roe2", "chk1", "chk2", "chk3") else label
        dup.write(r, 0, text, style)

    for name in ORDER:
        years = WINDOWS[name]
        for i, year in enumerate(years):
            c = col_of[(name, year)]
            is_base = i == 0
            if is_base:
                for r, _text, key in ratio_labels:
                    if key is None:
                        continue
                    dup.write_blank(r, c, None, base)
                continue
            prev = years[i - 1]

            def pair(key, _year=year, _prev=prev):
                return src_ref(name, _prev, key), src_ref(name, _year, key)

            ta0, ta1 = pair("ta")
            eq0, eq1 = pair("eq")
            rev = src_ref(name, year, "rev")
            ni = src_ref(name, year, "ni")
            oi = src_ref(name, year, "oi")
            nd0 = src_extra(name, prev, r_nd)
            nd1 = src_extra(name, year, r_nd)
            nopat = src_extra(name, year, r_nopat)
            nfe = src_extra(name, year, r_nfe)
            # Outer parentheses stop Excel from reading "a/b/2" as (a/b)/2.
            avg_ta = f"((({ta0})+({ta1}))/2)"
            avg_eq = f"((({eq0})+({eq1}))/2)"
            avg_nd = f"((({nd0})+({nd1}))/2)"
            avg_noa = f"({avg_eq}+{avg_nd})"
            # Row numbers on this sheet are 0-based in write(); Excel rows are +1.
            # Formulas refer to this sheet's own ratio cells where a check needs them.
            cl = xlsxwriter.utility.xl_col_to_name(c)
            dup.write_formula(6, c, f"={ni}/{rev}", pct_f)
            dup.write_formula(7, c, f"={rev}/{avg_ta}", x_f)
            dup.write_formula(8, c, f"={avg_ta}/{avg_eq}", x_f)
            dup.write_formula(9, c, f"={ni}/{avg_eq}", pct_b)
            dup.write_formula(10, c, f'=IF(ABS({cl}7*{cl}8*{cl}9-{cl}10)<0.0000005,"YES","NO")', yes_f)
            dup.write_formula(13, c, f"={nopat}/{rev}", pct_f)
            dup.write_formula(14, c, f"={rev}/{avg_noa}", x_f)
            dup.write_formula(15, c, f"={cl}14*{cl}15", pct_b)
            dup.write_formula(16, c, f'=IF(ABS({cl}14*{cl}15-{cl}16)<0.0000005,"YES","NO")', yes_f)
            dup.write_formula(18, c, f"={cl}16", pct_f)
            dup.write_formula(19, c, f'=IF(ABS({avg_nd})<250,"n.m.",{nfe}/({avg_nd}))', pct_f)
            dup.write_formula(20, c, f'=IF({cl}20="n.m.","n.m.",{cl}19-{cl}20)', pct_f)
            dup.write_formula(21, c, f"={avg_nd}/{avg_eq}", x_f)
            dup.write_formula(22, c, f'=IF({cl}21="n.m.","n.m.",{cl}21*{cl}22)', pct_b)
            dup.write_formula(23, c, f"={ni}/{avg_eq}", pct_b)
            dup.write_formula(24, c, f'=IF({cl}23="n.m.","n.m.",IF(ABS({cl}19+{cl}23-{cl}24)<0.0000005,"YES","NO"))', yes_f)
            dup.write_formula(26, c, f"=({oi})*(1-0.21)/({avg_noa})", pct_f)

    dup.write(28, 0, "NOPAT uses the statutory 21% federal rate on net interest only. Other income, gains, and the effective tax on operations stay inside net income, and therefore inside NOPAT. The last row taxes operating income at 21% and divides by the same average net operating assets, so the gap is the non-operating items and the tax-rate difference. Equity is the balance-sheet total. Intel's total includes non-controlling interests. Net income is consolidated net income, which for Intel is not the attributable-to-Intel line. Negative net financial leverage means net cash.", note)
    dup.set_row(28, 48)
    wb.close()


# ---------------------------------------------------------------------------
# Charts

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
    plt.rcParams["font.family"] = "Calibri"
    # Calendar position of each fiscal year-end. NVIDIA ends late January;
    # AMD and Intel end in late December, so NV FY t sits about a month after peer FY t-1.
    def xs(name):
        years = WINDOWS[name][1:]
        if name == "NVIDIA":
            return [y + 0.08 for y in years]
        return [(y + 1) - 0.04 for y in years]

    fig, axes = plt.subplots(1, 3, figsize=(7.35, 2.55), dpi=160)
    series = [
        ("Operating margin", "opm"),
        ("ROE", "roe"),
        ("Operating ROA", "rnoa"),
    ]
    for ax, (title, key) in zip(axes, series):
        for name in ORDER:
            ys = [RATIOS[name][i][key] * 100 for i in range(len(RATIOS[name]))]
            ax.plot(xs(name), ys, color=HEXES[name], marker="o", ms=3.5, lw=1.7, label=name, zorder=3)
        _style(ax)
        ax.set_title(title, fontsize=9, color="#1B2A4A", loc="left", pad=6)
        ax.set_xlim(2021.7, 2026.85)
        ax.set_xticks([2022, 2023, 2024, 2025, 2026])
        ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda v, _p: f"{v:.0f}%"))
    handles, labels = axes[0].get_legend_handles_labels()
    fig.legend(handles, labels, loc="upper center", ncol=3, frameon=False, fontsize=8)
    fig.tight_layout(rect=(0, 0, 1, 0.88))
    path = CHART / "performance.png"
    fig.savefig(path)
    plt.close()

    decomp_paths = []
    for name in ORDER:
        years = [y % 100 for y in WINDOWS[name][1:]]
        labels = [f"FY{y:02d}" for y in years]
        rnoa = [d["rnoa"] * 100 for d in RATIOS[name]]
        gain = [d["gain"] * 100 for d in RATIOS[name]]
        roe = [d["roe"] * 100 for d in RATIOS[name]]
        fig, ax = plt.subplots(figsize=(2.55, 2.45), dpi=160)
        x = list(range(len(labels)))
        width = 0.36
        ax.bar([i - width / 2 for i in x], rnoa, color=HEXES[name], width=width, zorder=2)
        ax.bar([i + width / 2 for i in x], gain, color="#8E9BAA", width=width, zorder=2)
        ax.plot(x, roe, color="#1B2A4A", marker="D", ms=4, lw=0, zorder=4)
        _style(ax)
        ax.set_xticks(x)
        ax.set_xticklabels(labels, fontsize=9, color="#1C2430")
        ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda v, _p: f"{v:.0f}%"))
        ax.set_title(name, fontsize=10, color=HEXES[name], loc="left")
        ax.tick_params(axis="y", labelsize=8)
        fig.subplots_adjust(left=0.18, right=0.97, top=0.88, bottom=0.16)
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
    nv, amd, intel = latest("NVIDIA"), latest("AMD"), latest("Intel")
    labels = [
        ("Latest full year", "FY26, ended Jan 25, 2026", "FY25, ended Dec 27, 2025", "FY25, ended Dec 27, 2025"),
        ("Revenue", bn(NV[2026]["rev"]), bn(AMD[2025]["rev"]), bn(INTC[2025]["rev"])),
        ("Operating margin", pct(nv["opm"]), pct(amd["opm"]), pct(intel["opm"])),
        ("ROE (average equity)", pct(nv["roe"]), pct(amd["roe"]), pct(intel["roe"], 2)),
        ("Operating ROA (modified DuPont)", pct(nv["rnoa"], 0), pct(amd["rnoa"]), pct(intel["rnoa"], 2)),
        ("Net cash (net debt) at year-end", f"Net cash {bn(abs(nv['nd']))}", f"Net cash {bn(abs(amd['nd']))}", f"Net debt {bn(intel['nd'])}"),
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
    widths = [2.15 * inch, 1.70 * inch, 1.70 * inch, 1.70 * inch]
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
    """One block per company, native fiscal years, five ratio years."""
    blocks = []
    metrics = [
        ("Revenue", lambda book, y, d: bn(book[y]["rev"])),
        ("Operating margin", lambda book, y, d: pct(d["opm"])),
        ("Net margin", lambda book, y, d: pct(d["npm"], 2 if abs(d["npm"]) < 0.005 else 1)),
        ("Capex / revenue", lambda book, y, d: pct(d["capex_rev"], 0)),
        ("Operating cash flow / capex", lambda book, y, d: f"{d['ocf_capex']:.1f}x"),
        ("Net debt (net cash) / equity", lambda book, y, d: pct(d["nd_eq"], 0)),
    ]
    for name in ORDER:
        years = WINDOWS[name][1:]
        header = [P(name, "thl")] + [P(f"FY{y % 100:02d}", "th") for y in years]
        data = [header]
        for label, fn in metrics:
            vals = []
            for i, year in enumerate(years):
                vals.append(P(fn(BOOKS[name], year, RATIOS[name][i]), "rightb" if label == "Revenue" else "right"))
            data.append([P(label, "tdb" if label == "Revenue" else "td")] + vals)
        color = COLORS[name]
        widths = [2.05 * inch] + [1.04 * inch] * 5
        table = Table(data, colWidths=widths)
        cmds = [
            ("BACKGROUND", (0, 0), (-1, 0), color),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 3),
            ("RIGHTPADDING", (0, 0), (-1, -1), 3),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ("LINEBELOW", (0, 1), (-1, -2), 0.25, RULE),
        ]
        for i in range(1, len(data)):
            if i % 2 == 0:
                cmds.append(("BACKGROUND", (0, i), (-1, i), ZEBRA))
        table.setStyle(TableStyle(cmds))
        blocks.append(KeepTogether([table, Spacer(1, 6)]))
    return blocks


def dupont_blocks():
    """One portrait table per company. The first column of years is the shaded base year."""
    label_w = 2.32 * inch
    year_w = 0.82 * inch
    widths = [label_w] + [year_w] * 6
    blocks = []
    yes_bg = colors.HexColor("#F7F1F2")
    section_bg = colors.HexColor("#F4F7F8")

    def value_row(name, label, key, kind, bold=False):
        row = [P(label, "tdb" if bold else "td")]
        row.append(P("", "right"))
        style = "rightb" if bold else "right"
        for i in range(len(WINDOWS[name]) - 1):
            figure = RATIOS[name][i][key]
            text = pct2(figure) if kind == "pct" else xn(figure)
            row.append(P(text, style))
        return row

    def yes_row(label):
        row = [P(label, "td"), P("", "yes")]
        row.extend(P("YES", "yes") for _ in range(5))
        return row

    for name in ORDER:
        years = WINDOWS[name]
        header = [P(name, "thl")]
        for i, year in enumerate(years):
            header.append(P(f"FY{year % 100:02d}", "thdark" if i == 0 else "th"))
        rows = [
            header,
            [P("Traditional DuPont", "tdb")] + [P("", "td")] * 6,
            value_row(name, "Net margin", "npm", "pct"),
            value_row(name, "× Asset turnover (sales / avg assets)", "ato", "x"),
            value_row(name, "× Leverage (avg assets / avg equity)", "em", "x"),
            value_row(name, "ROE (net income / avg equity)", "roe", "pct", True),
            yes_row("Check: ROE = margin × turnover × leverage"),
            [P("Modified DuPont", "tdb")] + [P("", "td")] * 6,
            value_row(name, "NOPAT margin", "pm", "pct"),
            value_row(name, "× Operating asset turnover (sales / avg NOA)", "turn", "x"),
            value_row(name, "Operating ROA", "rnoa", "pct", True),
            yes_row("Check: operating ROA = NOPAT margin × NOA turn"),
            value_row(name, "Operating ROA", "rnoa", "pct"),
            value_row(name, "− Net interest cost (after-tax interest / avg net debt)", "nbc", "pct"),
            value_row(name, "Spread", "spread", "pct"),
            value_row(name, "× Net financial leverage (avg net debt / avg equity)", "flev", "x"),
            value_row(name, "Gain or loss on financial leverage", "gain", "pct", True),
            value_row(name, "ROE (net income / avg equity)", "roe", "pct", True),
            yes_row("Check: ROE = operating ROA + gain"),
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
            ("BACKGROUND", (0, 17), (0, 17), yes_bg),
            ("BACKGROUND", (2, 17), (-1, 17), yes_bg),
            ("BACKGROUND", (1, 0), (1, -1), BASE_BG),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 3),
            ("RIGHTPADDING", (0, 0), (-1, -1), 3),
            ("TOPPADDING", (0, 0), (-1, -1), 2),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
            ("LINEBELOW", (0, 1), (-1, -2), 0.25, RULE),
            ("LINEABOVE", (0, -1), (-1, -1), 0.6, NAVY),
        ]
        table.setStyle(TableStyle(cmds))
        blocks.append(KeepTogether([table, Spacer(1, 8)]))
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
    story.append(P("Accounting, performance and strategy, with DuPont and modified DuPont for the last five fiscal years", "sub"))
    story.append(P("Prepared October 6, 2026  ·  Excel backup: NVIDIA_vs_AMD_Intel_DuPont_backup.xlsx", "sub"))
    story.append(Spacer(1, 8))
    story.append(snapshot_table())
    story.append(Spacer(1, 8))
    story.append(P("Bottom line", "h"))
    story.append(Spacer(1, 3))
    story.append(bullet(
        "Three different businesses, one customer.",
        f"NVIDIA designs the AI system and earned a {pct(nv['opm'])} operating margin in FY26. AMD designs the merchant alternative and earned {pct(amd['opm'])}. Intel builds its own wafers: the product groups earned $12.7bn of operating income and the foundry lost $10.3bn, so the company lost {bn(abs(INTC[2025]['oi']))} on operations.",
    ))
    story.append(bullet(
        "The year-ends are a month apart. The ROE gap is not a timing gap.",
        f"NVIDIA's FY26 ended January 25, 2026. AMD's and Intel's FY25 ended December 27, 2025. Average-equity ROE is {pct(nv['roe'])}, {pct(amd['roe'])}, and {pct(intel['roe'], 2)}.",
    ))
    story.append(bullet(
        "ROE is earned in operations. Cash, not debt, is what moves NVIDIA's.",
        f"Financial leverage adds or subtracts a few points at AMD and Intel. NVIDIA's net cash of {bn(abs(nv['nd']))} pulls ROE ({pct(nv['roe'])}) {pct(abs(nv['gain']), 0)} below its operating ROA ({pct(nv['rnoa'], 0)}).",
    ))
    story.append(bullet(
        "The accounting stories differ.",
        "NVIDIA's FY26 net income includes about $8.9bn of pretax equity-security gains, and goodwill rose $14.4bn for a Groq license and a group of employees. AMD's FY25 net income includes an $853m tax benefit. Intel's $26m of net income is a $5.6bn Altera gain and a valuation-allowance tax charge sitting on a $2.2bn operating loss.",
    ))
    story.append(bullet(
        "Fraud screen: nothing found that points to fraud.",
        "The items worth diligence are NVIDIA's new goodwill and investment gains, Intel's non-operating profit and capitalized interest, and AMD's tax release. Details are in the addendum.",
    ))
    story.append(Spacer(1, 4))
    story.append(P("How to read the periods", "h"))
    story.append(Spacer(1, 2))
    story.append(P("NVIDIA's fiscal year ends on the last Sunday in January. FY26 ended January 25, 2026. AMD and Intel end on the last Saturday in December. Their FY25 ended December 27, 2025. A NVIDIA fiscal year labeled t is the economic neighbor of the peers' fiscal year t−1, not of their year t. Peer FY26 is not in this filing set.", "body"))
    story.append(P("The ratio years are NVIDIA FY22–FY26 and AMD and Intel FY21–FY25. The shaded column before them is the base year used only for averages: NVIDIA FY21, AMD FY20, Intel FY20. There is no NVIDIA FY20 10-K in the set. FY21 supplies the opening balance.", "body"))
    story.append(P("DuPont tables follow the course slide. Balances are averages of the beginning and the end. Equity is the balance-sheet total, which includes non-controlling interests at Intel and is ordinary shareholders' equity at NVIDIA and AMD. Net income is income to the company: Intel's consolidated net income, not the line attributable to Intel. NOPAT adds back only after-tax net interest, at the 21% US statutory rate.", "body"))
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
    story.append(bullet("Profit in the latest year sits with the designer that does not own the fab.", f"Operating margin {pct(nv['opm'])} at NVIDIA, {pct(amd['opm'])} at AMD, {pct(intel['opm'])} at Intel."))

    story.append(section("2", "NVIDIA, AMD, and Intel side by side",
                         "The latest full year of each. NVIDIA FY26 sits about a month after the peers' FY25."))
    story.append(side_by_side())
    story.append(Spacer(1, 6))
    story.append(P("High-level differences", "h"))
    story.append(bullet("NVIDIA sells a system and is paid like a software company.", f"A {pct(nv['opm'])} operating margin on {bn(NV[2026]['rev'])} of revenue, with capex at {pct(nv['capex_rev'])} of revenue, is a designer's margin. The balance sheet is net cash."))
    story.append(bullet("AMD sells the alternative and still carries Xilinx.", "Goodwill and acquisition intangibles are 54% of assets. Operating margin is back to double digits. Turnover cannot look like NVIDIA's while that goodwill sits there."))
    story.append(bullet("Intel sells a turnaround that has not reached operating profit.", f"Revenue is down by a third from FY21. Capex is still {pct(intel['capex_rev'], 0)} of revenue. Net income of $26m is not evidence that the turnaround has arrived."))

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
    story.append(P("Comparability adjustment used in the ratios: none to the reported figures. The last row of the Excel shows the operating-income version of operating ROA beside the course-slide version. Section 5 discusses the gap.", "note"))

    story.append(section("5", "Performance, DuPont and modified DuPont",
                         "Five years. NVIDIA's trough is FY23. Intel's is FY24. AMD's is the year after Xilinx closed."))
    story.append(fitted_image(perf_chart, 7.25 * inch))
    story.append(P("Points are plotted at each fiscal year-end. NVIDIA ends in late January and the peers end in late December, so each NVIDIA point sits about a month to the right of the peers' comparable year. The latest points are NVIDIA FY26 and peer FY25.", "note"))
    story.extend(performance_table())
    story.append(P("Capex is cash spent on property and equipment, sign ignored. AMD's FY25 capex and operating cash flow include discontinued ZT operations, $38m and $1,216m. Intel's capex is the investing-section additions only. A further $3.0bn in 2025 and $1.2bn in 2024 are classified in financing. Net debt is debt plus long-term operating lease liabilities, minus cash and marketable or short-term investments. Negative means net cash.", "note"))
    story.append(bullet(
        "NVIDIA's revenue is a different shape.",
        f"From {bn(NV[2022]['rev'])} in FY22 to {bn(NV[2026]['rev'])} in FY26, 8.0 times. AMD went from {bn(AMD[2022]['rev'])} to {bn(AMD[2025]['rev'])}. Intel went from {bn(INTC[2021]['rev'])} in FY21 to {bn(INTC[2025]['rev'])}.",
    ))
    story.append(bullet(
        "The margin gap opened after FY23 and did not close.",
        f"NVIDIA's operating margin was {pct(nv23['opm'])} in the FY23 trough and {pct(nv['opm'])} in FY26. AMD's was {pct(row_of('AMD', 2023)['opm'])} in FY23 and {pct(amd['opm'])} in FY25. Intel's went from {pct(row_of('Intel', 2021)['opm'])} in FY21 to {pct(row_of('Intel', 2024)['opm'])} in FY24 and {pct(intel['opm'])} in FY25.",
    ))
    story.append(bullet(
        "Capital intensity says who owns the factory.",
        f"NVIDIA's capex was {pct(nv['capex_rev'])} of revenue and operating cash flow covered it {nv['ocf_capex']:.0f} times. AMD is in the same range. Intel's capex was {pct(row_of('Intel', 2023)['capex_rev'], 0)} of revenue in FY23 and still {pct(intel['capex_rev'], 0)} in FY25, and operating cash flow covered {intel['ocf_capex']:.2f}x of that spend.",
    ))
    story.append(bullet(
        "Balance sheets.",
        f"NVIDIA's net cash went from {bn(abs(row_of('NVIDIA', 2023)['nd']))} at the end of FY23 to {bn(abs(nv['nd']))}. AMD stayed in net cash and ended at {bn(abs(amd['nd']))}. Intel's net debt peaked at {bn(row_of('Intel', 2024)['nd'])} ({pct(row_of('Intel', 2024)['nd_eq'], 0)} of equity) in FY24 and fell to {bn(intel['nd'])} after the Altera proceeds and lower capex.",
    ))

    blocks = dupont_blocks()
    story.append(KeepTogether([
        Spacer(1, 6),
        P("Advanced DuPont decomposition of return on equity", "h"),
        P("Each company is on its own. The shaded column is the base year, used only for the averages. YES is the unrounded identity.", "deck"),
        blocks[0],
    ]))
    story.extend(blocks[1:])
    story.append(P("Layout and definitions follow the course slide. Equity is the average of beginning and ending balance-sheet equity. Intel's total includes non-controlling interests. Net income is income to the company, so Intel uses consolidated net income rather than net income attributable to Intel. NOPAT = net income + (interest expense − interest income) × (1 − 21%). Net debt = short-term debt, the current portion, and long-term debt, plus the long-term operating lease liability where it is its own balance-sheet line, minus cash and marketable securities or short-term investments. Negative net financial leverage means net cash. Shaded columns are the base year. YES is the unrounded identity; the printed figures are rounded to two decimals. Current operating lease liabilities sit inside accrued liabilities and are not in net debt. Intel's operating leases, about $0.4bn, are in other liabilities and are not in net debt. Intel's FY20 trading assets of $15.7bn are equity securities and are not treated as liquid investments. NVIDIA's $22.3bn of non-marketable equity securities stay inside net operating assets.", "note"))

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
    story.append(P("Each panel is that company's own fiscal years. The colored bar is operating ROA. The gray bar is the gain or loss from financial leverage. Diamonds are ROE. The scales differ.", "note"))
    story.append(bullet(
        "ROE is an operating story at all three.",
        f"AMD's leverage effect in FY25 is {pct(amd['gain'])}. Intel's is {pct(intel['gain'])}. Neither explains the gap with NVIDIA. The entire NVIDIA gap versus its own operating ROA is the cash drag of {pct(nv['gain'])}.",
    ))
    story.append(bullet(
        "At NVIDIA, margin sets the level and the asset base sets the latest move.",
        f"Net margin stayed near {pct(row_of('NVIDIA', 2025)['npm'])} and then {pct(nv['npm'])}. Operating ROA fell from {pct(row_of('NVIDIA', 2025)['rnoa'], 0)} to {pct(nv['rnoa'], 0)} because sales grew more slowly than net operating assets. Ending net operating assets are equity plus net debt, {bn(NV[2026]['eq'] + nv['nd'])}. Inside that number are $22.3bn of non-marketable equity securities, $21.4bn of inventory, and the Groq goodwill. The product margin did not fall. The balance sheet got heavier.",
    ))
    story.append(bullet(
        "AMD's ROE broke when Xilinx arrived, and it is the goodwill that is still holding it down.",
        f"FY21 ROE was {pct(row_of('AMD', 2021)['roe'])} on a small asset base. FY22 ROE was {pct(row_of('AMD', 2022)['roe'])} after assets jumped from {bn(AMD[2021]['ta'])} to {bn(AMD[2022]['ta'])}. FY25 operating ROA of {pct(amd['rnoa'])} is a {pct(amd['opm'])} margin on 0.47x total-asset turnover. Taxing operating income at 21% instead of using course-slide NOPAT, operating ROA would be {pct(amd['alt'])}. The difference is the tax benefit and the small net-interest add-back.",
    ))
    story.append(bullet(
        "Intel's operating ROA near zero is not a sign of a healed business.",
        f"Course-slide NOPAT keeps the Altera gain and the valuation-allowance charge. Operating ROA is {pct(intel['rnoa'], 2)}. Using operating income taxed at 21%, it is {pct(intel['alt'])}. In FY24 the same two measures were {pct(row_of('Intel', 2024)['rnoa'])} and {pct(row_of('Intel', 2024)['alt'])}: both losses, and the course-slide number was worse because of the impairment and the tax charge. Debt is not the reason. Net financial leverage was {xn(intel['flev'])} in FY25.",
    ))
    story.append(bullet(
        "Reported borrowing cost understates Intel's build-out.",
        f"FY25 after-tax net interest over average net debt is a borrowing cost of {pct(intel['nbc'])}. Interest expense of $1,091m is after $1.2bn capitalized into the fabs. In FY22, interest income exceeded the reported interest expense while the company was in net debt, so net borrowing cost was negative. The cash-flow statement shows cash interest, net of what was capitalized, of $1,106m in 2025.",
    ))
    story.append(bullet(
        "Idle cash is NVIDIA's ROE question, the same way a foundry loss is Intel's.",
        f"NVIDIA's net cash earns a low single-digit yield. The modified-DuPont borrowing cost on that net cash position is {pct(nv['nbc'])}, against an operating ROA of {pct(nv['rnoa'], 0)}. Buybacks of $40.4bn and the $17.5bn of equity-security purchases are the two uses already visible. How fast the cash is returned, spent on the $22.7bn of leases not yet started, or left invested, will move ROE more than any borrowing will.",
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

    story.append(CondPageBreak(3.2 * inch))
    story.append(Spacer(1, 8))
    story.append(P("Sources and method", "h"))
    story.append(Spacer(1, 3))
    story.append(P("Figures", "h"))
    story.append(P("Every ratio input is a line from the company's Form 10-K in this project's filing folder. Amounts are USD millions. The Excel source sheet holds the lines. The DuPont sheet computes every ratio, and the YES cells test the identity.", "body"))
    story.append(P("NVIDIA FY22–FY26 income statement, cash, marketable securities, debt, equity, assets, operating cash flow, and capex are the statement lines in the FY26, FY24, and FY23 Form 10-Ks. FY21 is the current-year column of the FY21 Form 10-K. Long-term operating lease liabilities are the balance-sheet line. The FY24 securities balance of $18,704m is the FY24 balance sheet.", "body"))
    story.append(P("AMD FY20–FY25 income statement, cash, short-term investments, debt, leases, equity, and assets are the statement lines in the FY20, FY22, FY24, and FY25 Form 10-Ks. Interest income for FY20–FY22 is the note table in the FY22 Form 10-K ($8m, $8m, $65m). Interest income for FY23–FY25 is the note table in the FY25 Form 10-K ($206m, $182m, $215m). FY25 capex of $1,012m is purchases of property and equipment of $974m plus $38m in discontinued operations. FY25 operating cash flow of $7,709m is $6,493m continuing plus $1,216m discontinued.", "body"))
    story.append(P("Intel FY20–FY25 revenue, operating income, cash, short-term investments, debt, equity, and assets are the statement lines. Consolidated net income is the total, not the attributable line: FY25 $26m, FY24 a loss of $19,233m, FY23 $1,675m, FY22 $8,017m. Interest income and interest expense are the interest-and-other note: FY20 from the FY20 10-K, FY21–FY23 from the FY23 10-K, FY24–FY25 from the FY25 10-K. Debt equals the note total (short-term plus long-term), which matches the selected-data debt figure in FY20 ($36,401m). Operating cash flow for FY20 and FY21 is the revised comparative in the FY23 and FY24 cash-flow statements. Capex is additions to property, plant and equipment in the investing section.", "body"))
    story.append(P("Context used in the prose, each from the latest 10-K: NVIDIA Groq note, segment revenue, customer concentration, geographic revenue, inventory provisions, lease commitments, buybacks, and the foundry and memory supplier list. AMD segment results, TSMC dependence, China revenue, the tax reconciliation, buybacks, and goodwill. Intel segment operating income, the Altera gain, capitalized interest, partner contributions, customer concentration, China billings, and the financing-section equipment additions.", "body"))
    story.append(P("Limits", "h"))
    story.append(P("The latest columns are one month apart, not one year. Pairing NVIDIA FY26 with AMD or Intel FY26 would be wrong, because those peer years are not filed here.", "body"))
    story.append(P("Net debt leaves out current operating lease liabilities, which are inside accrued liabilities, and leaves out Intel's operating leases of about $0.4bn. It leaves out Intel's FY20 trading assets of $15.7bn because they are equity securities rather than cash or debt securities. It leaves NVIDIA's non-marketable equity securities inside net operating assets. Moving any of those choices would change operating ROA. The direction of the comparison would not change.", "body"))
    story.append(P("Intel's second equipment line, classified in financing, is disclosed and is not in the capex ratio. Adding it would make FY25 capex $17.7bn and the cash-flow coverage weaker.", "body"))
    story.append(P("Audit fees, headcount, and a full five-year auditor-tenure check were not done. Market prices and the most recent quarter are not in this annual-report set, so they are not in the snapshot.", "body"))
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
    # Sanity: snapshot numbers match the decomposition.
    nv = latest("NVIDIA")
    assert abs(nv["roe"] - (nv["rnoa"] + nv["gain"])) < 1e-9
    print("wrote", OUT)
    print("wrote", XLSX)
    print("NVIDIA FY26 ROE", pct(nv["roe"]), "RNOA", pct(nv["rnoa"]), "gain", pct(nv["gain"]))
    print("AMD FY25 ROE", pct(latest("AMD")["roe"]), "Intel FY25 ROE", pct(latest("Intel")["roe"], 2))


if __name__ == "__main__":
    main()
