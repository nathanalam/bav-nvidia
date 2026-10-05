#!/usr/bin/env python3
"""Corrected BAV Deliverable #3: NVIDIA accounting and performance."""

import os
from pathlib import Path
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY, TA_RIGHT
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image,
    KeepTogether, ListFlowable, ListItem, CondPageBreak,
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

def _font(*names):
    roots = [Path("/usr/share/fonts"), Path("/Library/Fonts"), Path.home() / "Library/Fonts", Path("C:/Windows/Fonts")]
    for root in roots:
        if not root.exists():
            continue
        for name in names:
            hits = list(root.rglob(name))
            if hits:
                return str(hits[0])
    raise FileNotFoundError(names[0])

pdfmetrics.registerFont(TTFont("Arial", _font("LiberationSans-Regular.ttf", "Arial.ttf")))
pdfmetrics.registerFont(TTFont("Arial-Bold", _font("LiberationSans-Bold.ttf", "Arial Bold.ttf")))
pdfmetrics.registerFont(TTFont("Arial-Italic", _font("LiberationSans-Italic.ttf", "Arial Italic.ttf")))
pdfmetrics.registerFont(TTFont("Arial-BoldItalic", _font("LiberationSans-BoldItalic.ttf", "Arial Bold Italic.ttf")))

OUT = str(Path(__file__).resolve().parent / "reports" / "Alam_NVIDIA_Accounting_Performance.pdf")
CHART = str(Path(__file__).resolve().parent / "reports" / "charts")
Path(OUT).parent.mkdir(parents=True, exist_ok=True)
os.makedirs(CHART, exist_ok=True)

NAVY = colors.HexColor("#1B3A4B")
TEAL = colors.HexColor("#1F4E5F")
RULE = colors.HexColor("#D0D5DD")
HEADBG = colors.HexColor("#1B3A4B")
ALT = colors.HexColor("#F4F7F8")
GOLD = colors.HexColor("#F7F1E8")
GREEN = "#1F7A4D"
RED = "#B42318"
BLUE = "#1D4E89"
INK = colors.HexColor("#1C1814")

# Filing extracts (USD millions). NVIDIA FY ends late January.
# Income-statement sheet in the workbook has shifted percent columns; NI, revenue,
# and operating margin are taken from the curated DuPont series and cash-flow extract.
NV = {
    2023: dict(rev=26974, ni=4368, oi=4224, ta=41182, eq=22101, cash=3389, sec=9907, debt=1250+9703, ocf=5641, capex=1833, sbc=2709),
    2024: dict(rev=60922, ni=29760, oi=32972, ta=65728, eq=42978, cash=7280, sec=18704, debt=1250+8459, ocf=28090, capex=1069, sbc=3549),
    2025: dict(rev=130497, ni=72880, oi=81453, ta=111601, eq=79327, cash=8589, sec=34621, debt=0+8463, ocf=64089, capex=3236, sbc=4737),
    2026: dict(rev=215938, ni=120067, oi=130427, ta=206803, eq=157293, cash=10605, sec=51951, debt=999+7469, ocf=102718, capex=6042, sbc=6386,
              ar=38466, inv=21403, gw=20832, tax=21383, ebt=141450),
}
# OI FY24 = 54.1% of revenue; FY26 OI = 60.4% of revenue (workbook percent columns).
NV[2024]["oi"] = round(60922 * 0.541)
NV[2026]["oi"] = round(215938 * 0.604)

AMD = {
    2022: dict(rev=23601, ni=1320, ta=67580, eq=54750, oi=1264),
    2023: dict(rev=22680, ni=854, ta=67885, eq=55892, oi=401),
    2024: dict(rev=25785, ni=1641, ta=69226, eq=57568, oi=1900),
    2025: dict(rev=34639, ni=4335, ta=76926, eq=62999, oi=3694),
}
INTC = {
    2022: dict(rev=63054, ni=8017, ta=182103, eq=103286, oi=2334),
    2023: dict(rev=54228, ni=1675, ta=191572, eq=109965, oi=93),
    2024: dict(rev=53101, ni=-19233, ta=196485, eq=105032, oi=-11678),
    2025: dict(rev=52853, ni=-267, ta=211429, eq=126360, oi=-2214),
}

# Advanced DuPont from curated average-balance decomposition (Palepu / Healy).
# Tuple: rnoa, nbc, spread, flev, roe_decomp, roe_actual
ADV = {
    "NVIDIA": {2023: (17.3, -4.4, 21.6, 0.0, 17.4, 17.9),
               2024: (113.1, 7.8, 105.4, -0.21, 90.8, 91.5),
               2025: (168.9, 6.9, 162.0, -0.32, 117.7, 119.2),
               2026: (149.8, 3.9, 145.9, -0.38, 95.0, 101.5)},
    "AMD": {2022: (4.4, -3.6, 8.0, -0.08, 3.8, 4.2),
            2023: (0.8, -4.6, 5.4, -0.04, 0.5, 1.5),
            2024: (3.0, -2.8, 5.9, -0.05, 2.8, 2.9),
            2025: (6.8, -2.3, 9.0, -0.10, 5.9, 7.2)},
    "Intel": {2022: (2.3, 0.0, 2.3, 0.02, 2.3, 8.1),
              2023: (0.1, 0.0, 0.1, 0.03, 0.1, 1.6),
              2024: (-9.6, 0.0, -9.6, -0.04, -9.2, -17.9),
              2025: (-2.1, 0.0, -2.1, -0.22, -1.6, -0.2)},
}

# Calendar alignment: NVIDIA fiscal year ends late January, so NV FY t matches peer FY t-1.
ALIGNED = [("YE23", 2023, 2022), ("YE24", 2024, 2023), ("YE25", 2025, 2024), ("YE26", 2026, 2025)]
LABELS = [a[0] for a in ALIGNED]


def pct(n, d):
    return n / d


def trad(row):
    npm = row["ni"] / row["rev"]
    ato = row["rev"] / row["ta"]
    em = row["ta"] / row["eq"]
    roa = row["ni"] / row["ta"]
    roe = row["ni"] / row["eq"]
    return npm, ato, em, roa, roe


def advanced(oi, ni, equity, debt, fin, tax):
    """Ending-balance Palepu identity. NFE is the plug so NOPAT - NFE = NI."""
    nopat = oi * (1 - tax)
    nfe = nopat - ni
    nfo = debt - fin
    noa = equity + nfo
    rnoa = nopat / noa
    flev = nfo / equity
    nbc = nfe / nfo
    roe = ni / equity
    decomp = rnoa + (rnoa - nbc) * flev
    return dict(nopat=nopat, nfe=nfe, nfo=nfo, noa=noa, rnoa=rnoa, nbc=nbc,
                spread=rnoa - nbc, flev=flev, effect=(rnoa - nbc) * flev,
                roe=roe, decomp=decomp)


NV_FIN = {
    2023: dict(debt=1250 + 9703, fin=3389 + 9907, tax=0.0),
    2024: dict(debt=1250 + 8459, fin=7280 + 18704, tax=0.12),
    2025: dict(debt=8463, fin=8589 + 34621, tax=11146 / 84026),
    2026: dict(debt=999 + 7469, fin=10605 + 51951, tax=21383 / 141450),
}
AMD_FIN = {
    2022: dict(debt=2467, fin=4835 + 1020, tax=0.0),
    2023: dict(debt=751 + 1717, fin=3933 + 1840, tax=0.0),
    2024: dict(debt=1721, fin=3787 + 1345, tax=0.14),
    2025: dict(debt=874 + 2348, fin=5539 + 5013, tax=0.0),
}
INTC_FIN = {
    2022: dict(debt=4367 + 41682, fin=11144 + 17194, tax=0.0),
    2023: dict(debt=2288 + 46978, fin=7079 + 17955, tax=0.0),
    2024: dict(debt=3729 + 46282, fin=8249 + 13813, tax=0.0),
    2025: dict(debt=2499 + 44086, fin=14265 + 23151, tax=0.0),
}


def run_forecast(rev0, bv0, growth, margins, re=0.12, g=0.035, tax=0.15):
    rev, bv = rev0, bv0
    revs, nis, ris, pvs = [], [], [], []
    for t, (gi, om) in enumerate(zip(growth, margins), start=1):
        rev *= (1 + gi)
        ni = rev * om * (1 - tax)
        ri = ni - re * bv
        revs.append(rev)
        nis.append(ni)
        ris.append(ri)
        pvs.append(ri / ((1 + re) ** t))
        bv += ni
    pv_ri = sum(pvs)
    pv_tv = ris[-1] * (1 + g) / (re - g) / ((1 + re) ** 5)
    return dict(revs=revs, nis=nis, ris=ris, pvs=pvs, pv_ri=pv_ri, pv_tv=pv_tv,
                equity=bv0 + pv_ri + pv_tv, b0=bv0)


FORECASTS = {
    "NVIDIA": run_forecast(215938, 157293, [0.25, 0.20, 0.14, 0.10, 0.07], [0.58, 0.55, 0.52, 0.50, 0.48]),
    "AMD": run_forecast(34639, 62999, [0.25, 0.20, 0.16, 0.13, 0.10], [0.14, 0.17, 0.19, 0.21, 0.22]),
    "Intel": run_forecast(52853, 126360, [0.02, 0.04, 0.06, 0.05, 0.04], [0.02, 0.06, 0.09, 0.11, 0.12]),
}
SHARES = {"NVIDIA": 24514, "AMD": 1636, "Intel": 4530}
PRICES = {"NVIDIA": 236, "AMD": 632, "Intel": 117}


def style_ax(ax):
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_color("#D0D5DD")
    ax.spines["bottom"].set_color("#D0D5DD")
    ax.tick_params(colors="#3D3A36", labelsize=8)
    ax.grid(axis="y", color="#EEF1F4", zorder=0)
    ax.set_axisbelow(True)


def make_charts():
    plt.rcParams["font.family"] = "Liberation Sans"
    x = list(range(4))
    nv_roe = [NV[n]["ni"] / NV[n]["eq"] * 100 for _, n, _ in ALIGNED]
    amd_roe = [AMD[p]["ni"] / AMD[p]["eq"] * 100 for _, _, p in ALIGNED]
    int_roe = [INTC[p]["ni"] / INTC[p]["eq"] * 100 for _, _, p in ALIGNED]

    fig, ax = plt.subplots(figsize=(7.2, 3.15), dpi=160)
    ax.plot(x, nv_roe, color=GREEN, marker="o", lw=2.2, label="NVIDIA")
    ax.plot(x, amd_roe, color=RED, marker="s", lw=1.8, label="AMD")
    ax.plot(x, int_roe, color=BLUE, marker="^", lw=1.8, label="Intel")
    ax.axhline(0, color="#98A2B3", lw=0.6)
    style_ax(ax)
    ax.set_ylabel("ROE, ending equity (%)", fontsize=8)
    ax.set_xticks(x)
    ax.set_xticklabels(LABELS)
    ax.legend(frameon=False, fontsize=8, loc="upper left")
    ax.set_title("ROE, aligned year-ends (NVIDIA FY t = peer FY t−1)", fontsize=10, color="#1B3A4B", loc="left")
    fig.tight_layout()
    fig.savefig(f"{CHART}/roe.png")
    plt.close()

    fig, axes = plt.subplots(1, 3, figsize=(7.2, 2.85), dpi=160)
    series = [
        ("Net margin (%)", lambda d, y: d[y]["ni"] / d[y]["rev"] * 100),
        ("Asset turnover (x)", lambda d, y: d[y]["rev"] / d[y]["ta"]),
        ("Equity multiplier (x)", lambda d, y: d[y]["ta"] / d[y]["eq"]),
    ]
    width = 0.24
    for ax, (title, fn) in zip(axes, series):
        nv = [fn(NV, n) for _, n, _ in ALIGNED]
        amd = [fn(AMD, p) for _, _, p in ALIGNED]
        intel = [fn(INTC, p) for _, _, p in ALIGNED]
        ax.bar([i - width for i in x], nv, width, color=GREEN, label="NVIDIA")
        ax.bar(x, amd, width, color=RED, label="AMD")
        ax.bar([i + width for i in x], intel, width, color=BLUE, label="Intel")
        style_ax(ax)
        ax.set_title(title, fontsize=9, color="#1B3A4B")
        ax.set_xticks(x)
        ax.set_xticklabels(LABELS, fontsize=7)
        ax.axhline(0, color="#98A2B3", lw=0.5)
    axes[0].legend(frameon=False, fontsize=7, loc="upper left")
    fig.tight_layout()
    fig.savefig(f"{CHART}/drivers.png")
    plt.close()

    fig, ax = plt.subplots(figsize=(7.2, 3.15), dpi=160)
    ax.plot(x, [advanced(NV[n]["oi"], NV[n]["ni"], NV[n]["eq"], **NV_FIN[n])["rnoa"] * 100 for _, n, _ in ALIGNED], color=GREEN, marker="o", lw=2.2, label="NVIDIA")
    ax.plot(x, [advanced(AMD[p]["oi"], AMD[p]["ni"], AMD[p]["eq"], **AMD_FIN[p])["rnoa"] * 100 for _, _, p in ALIGNED], color=RED, marker="s", lw=1.8, label="AMD")
    ax.plot(x, [advanced(INTC[p]["oi"], INTC[p]["ni"], INTC[p]["eq"], **INTC_FIN[p])["rnoa"] * 100 for _, _, p in ALIGNED], color=BLUE, marker="^", lw=1.8, label="Intel")
    ax.axhline(0, color="#98A2B3", lw=0.6)
    style_ax(ax)
    ax.set_ylabel("RNOA (%)", fontsize=8)
    ax.set_xticks(x)
    ax.set_xticklabels(LABELS)
    ax.legend(frameon=False, fontsize=8)
    ax.set_title("RNOA, same aligned year-ends", fontsize=10, color="#1B3A4B", loc="left")
    fig.tight_layout()
    fig.savefig(f"{CHART}/rnoa.png")
    plt.close()

    nv_rev = [NV[n]["rev"] / 1000 for _, n, _ in ALIGNED]
    amd_rev = [AMD[p]["rev"] / 1000 for _, _, p in ALIGNED]
    int_rev = [INTC[p]["rev"] / 1000 for _, _, p in ALIGNED]
    fig, ax = plt.subplots(figsize=(7.2, 3.15), dpi=160)
    ax.plot(x, nv_rev, color=GREEN, marker="o", lw=2.2, label="NVIDIA")
    ax.plot(x, amd_rev, color=RED, marker="s", lw=1.8, label="AMD")
    ax.plot(x, int_rev, color=BLUE, marker="^", lw=1.8, label="Intel")
    fx = [3, 4, 5, 6, 7, 8]
    ax.plot(fx, [nv_rev[-1]] + [v / 1000 for v in FORECASTS["NVIDIA"]["revs"]], color=GREEN, lw=1.4, ls="--")
    ax.plot(fx, [amd_rev[-1]] + [v / 1000 for v in FORECASTS["AMD"]["revs"]], color=RED, lw=1.4, ls="--")
    ax.plot(fx, [int_rev[-1]] + [v / 1000 for v in FORECASTS["Intel"]["revs"]], color=BLUE, lw=1.4, ls="--")
    style_ax(ax)
    ax.set_ylabel("Revenue ($ billions)", fontsize=8)
    ax.set_xticks(list(range(9)))
    ax.set_xticklabels(LABELS + ["+1", "+2", "+3", "+4", "+5"], fontsize=7)
    ax.legend(frameon=False, fontsize=8)
    ax.set_title("Revenue, aligned history and five-year forecast", fontsize=10, color="#1B3A4B", loc="left")
    fig.tight_layout()
    fig.savefig(f"{CHART}/revenue.png")
    plt.close()
    return FORECASTS["NVIDIA"]["revs"]


FWD = make_charts()


def header_footer(canvas, doc):
    canvas.saveState()
    canvas.setFillColor(NAVY)
    canvas.rect(0, letter[1] - 28, letter[0], 28, fill=1, stroke=0)
    canvas.setFillColor(colors.white)
    canvas.setFont("Arial", 8)
    canvas.drawString(0.7 * inch, letter[1] - 18, "BAV  |  Accounting and Performance  |  NVIDIA Corporation")
    canvas.setFont("Arial", 8)
    canvas.drawRightString(letter[0] - 0.7 * inch, letter[1] - 18, "Nathan Shiham Alam")
    canvas.setFillColor(colors.HexColor("#667085"))
    canvas.setFont("Arial", 8)
    canvas.drawString(0.7 * inch, 0.38 * inch, "Harvard Business School  ·  Fall 2026  ·  Draft for Deliverable #3")
    canvas.drawRightString(letter[0] - 0.7 * inch, 0.38 * inch, f"{doc.page}")
    canvas.setStrokeColor(RULE)
    canvas.line(0.7 * inch, 0.52 * inch, letter[0] - 0.7 * inch, 0.52 * inch)
    canvas.restoreState()


styles = getSampleStyleSheet()
styles.add(ParagraphStyle(name="CoverTitle", fontName="Arial-Bold", fontSize=18, leading=22, textColor=NAVY, spaceAfter=2))
styles.add(ParagraphStyle(name="CoverSub", fontName="Arial", fontSize=11, leading=14, textColor=TEAL, spaceAfter=2))
styles.add(ParagraphStyle(name="Meta", fontName="Arial", fontSize=9, leading=12, textColor=colors.HexColor("#3D3A36"), spaceAfter=1))
styles.add(ParagraphStyle(name="H1", fontName="Arial-Bold", fontSize=12, leading=15, textColor=NAVY, spaceBefore=11, spaceAfter=4))
styles.add(ParagraphStyle(name="H2", fontName="Arial-Bold", fontSize=10.5, leading=13, textColor=TEAL, spaceBefore=8, spaceAfter=3))
styles.add(ParagraphStyle(name="Body", fontName="Arial", fontSize=9, leading=12.2, textColor=INK, alignment=TA_JUSTIFY, spaceAfter=6))
styles.add(ParagraphStyle(name="Note", fontName="Arial-Italic", fontSize=8, leading=10.4, textColor=colors.HexColor("#3D3A36"), spaceBefore=1, spaceAfter=6))
styles.add(ParagraphStyle(name="Caption", fontName="Arial-Italic", fontSize=8, leading=10, textColor=colors.HexColor("#3D3A36"), spaceBefore=1, spaceAfter=8))
styles.add(ParagraphStyle(name="Th", fontName="Arial-Bold", fontSize=7.5, leading=9.2, textColor=colors.white, alignment=TA_CENTER))
styles.add(ParagraphStyle(name="Td", fontName="Arial", fontSize=7.5, leading=9.2, textColor=INK, alignment=TA_RIGHT))
styles.add(ParagraphStyle(name="TdL", fontName="Arial", fontSize=7.5, leading=9.2, textColor=INK, alignment=TA_LEFT))
styles.add(ParagraphStyle(name="TdB", fontName="Arial-Bold", fontSize=7.5, leading=9.2, textColor=INK, alignment=TA_RIGHT))
styles.add(ParagraphStyle(name="TdLB", fontName="Arial-Bold", fontSize=7.5, leading=9.2, textColor=INK, alignment=TA_LEFT))
styles.add(ParagraphStyle(name="Formula", fontName="Arial", fontSize=8, leading=10.5, textColor=NAVY, alignment=TA_LEFT, spaceBefore=1, spaceAfter=4))


def P(text, style="Body"):
    return Paragraph(text, styles[style])


def th(text):
    return Paragraph(text, styles["Th"])


def td(text, bold=False, left=False):
    if left and bold:
        return Paragraph(text, styles["TdLB"])
    if left:
        return Paragraph(text, styles["TdL"])
    if bold:
        return Paragraph(text, styles["TdB"])
    return Paragraph(text, styles["Td"])


def table(rows, widths, highlight_last=False):
    sty = [
        ("BACKGROUND", (0, 0), (-1, 0), HEADBG),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Arial-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 7.5),
        ("ALIGN", (1, 1), (-1, -1), "RIGHT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("GRID", (0, 0), (-1, -1), 0.25, RULE),
        ("BACKGROUND", (0, 1), (-1, -1), colors.white),
    ]
    for i in range(1, len(rows)):
        if i % 2 == 0:
            sty.append(("BACKGROUND", (0, i), (-1, i), ALT))
    if highlight_last:
        sty.append(("BACKGROUND", (0, -1), (-1, -1), GOLD))
    t = Table(rows, colWidths=widths, repeatRows=1)
    t.setStyle(TableStyle(sty))
    return t


def fmt_pct(x):
    return f"{x*100:.1f}%"


def fmt_x(x):
    return f"{x:.2f}x"


story = []
story.append(Spacer(1, 8))
story.append(P("NVIDIA Corporation (NVDA)", "CoverTitle"))
story.append(P("Accounting and Performance Analysis", "CoverSub"))
story.append(Spacer(1, 4))
story.append(P("Nathan Shiham Alam", "Meta"))
story.append(P("Business Analysis and Evaluation  ·  Harvard Business School  ·  Fall 2026", "Meta"))
story.append(P("Individual draft for Deliverable #3  ·  October 5, 2026", "Meta"))
story.append(Spacer(1, 6))

story.append(P("1. Executive Summary and Recommendation", "H1"))
story.append(P(
    "NVIDIA is the standard platform for accelerated computing. CUDA, full-rack AI systems, and hyperscaler distribution produced FY2026 revenue of $215.9 billion, a 60.4% operating margin, and $96.7 billion of free cash flow. Ending-equity ROE is 76.3% on both decompositions. The advanced split is a 107.3% return on net operating assets minus a 30.9-point drag from negative financial leverage: NVIDIA is a net creditor.",
))
story.append(P(
    "The prior draft’s $280 target does not follow from the residual-income model. That figure required a $2.0 trillion off-balance-sheet plug. The filing-based model below, with no plug, is about $50 a share. A separate scenario file that starts from $303 billion of trailing revenue, rather than the $216 billion fiscal year, spans $69 to $272. The October 5, 2026 price of $236 is inside that span only on the aggressive case. Recommendation for this draft: <b>Hold</b>.",
))

story.append(P("2. Company and Operations Overview", "H1"))
story.append(P("2.1 Business", "H2"))
story.append(P(
    "Founded in 1993, NVIDIA moved from PC graphics to data-center AI infrastructure. It sells full-stack systems: GPUs, CPUs, networking, and software. End markets are Data Center, Gaming (GeForce), Professional Visualization (RTX and Omniverse), and Automotive (DRIVE). Data Center is the growth and profit driver. NVIDIA’s fiscal year ends in late January, so FY2026 is the year ended January 2026. AMD and Intel report on a December year-end. Comparisons below use aligned year-ends: NVIDIA FY t is paired with peer FY t−1. YE26 is NVIDIA FY2026 and AMD/Intel FY2025.",
))
story.append(P("2.2 Strategy", "H2"))
story.append(P(
    "The strategy is to make accelerated computing the default and to own the layer customers cannot easily leave. Three choices do the work. First, CUDA and the software stack raise switching costs: models, libraries, and engineers are written to NVIDIA, not to a chip. Second, the product is a system, not a card — Grace Blackwell and the Rubin roadmap are sold as AI-factory racks, which pulls networking and software into the same purchase. Third, distribution runs through hyperscalers and OEMs, so NVIDIA captures the training and inference build-out without owning the data center.",
))
story.append(P(
    "The financial signature of that strategy is a software-like margin on a hardware revenue base, with asset turnover near 1.0x rather than the sub-0.5x of a foundry. NVIDIA does not manufacture wafers. It designs, and it keeps the scarce resource — software compatibility and system integration — in-house.",
))
story.append(P("2.3 Risks", "H2"))
story.append(P(
    "Customer concentration is high: the largest customer was 22% of revenue, and two customers were 36%. Manufacturing is single-sourced in practice at TSMC, with Samsung as a limited second source. Hyperscalers are designing their own accelerators, and AMD is the merchant alternative. Export controls cap the China mix. None of these shows up as distress on the balance sheet — Altman Z is far above the gray zone — but they cap how long a 55% net margin can persist.",
))

story.append(P("3. Accounting Analysis", "H1"))
story.append(P(
    "Policies are generally conservative, and the cash-flow statement confirms earnings. Three items matter for earnings quality.",
))
story.append(P(
    "<b>Revenue and returns.</b> Product revenue is recognized when control transfers, typically on delivery to OEMs, distributors, and cloud customers. Sales-return allowances have scaled with revenue rather than being used to smooth growth. That is the right direction. The risk is not channel stuffing; it is customer concentration. A delay at two buyers moves the quarter.",
))
story.append(P(
    "<b>Stock-based compensation.</b> SBC was $6.4 billion in FY2026, up from $2.7 billion in FY2023, and is a real cost. Historical non-GAAP results excluded it, which overstated economic earnings relative to GAAP. Beginning in Q1 FY2027, non-GAAP metrics will include SBC. That change improves comparability and should be treated as the cleaner earnings number going forward. GAAP net income of $120.1 billion already includes SBC.",
))
story.append(P(
    "<b>Working capital, tax, and goodwill.</b> Receivables rose to $38.5 billion and inventories to $21.4 billion as Blackwell ramped. Inventory is carried at the lower of cost or net realizable value; a demand air pocket would hit gross margin through reserves, not through a hidden accrual. The effective tax rate on FY2026 pretax income of $141.5 billion was about 15%, and the deferred-tax valuation allowance declined to $768 million, so tax expense is no longer being shielded by a large allowance build. Goodwill stepped up to $20.8 billion from $5.2 billion. It is not amortized. If the acquired intangibles do not earn their keep, the impairment hits earnings in one period and will not have been previewed by the margin trend.",
))
story.append(P(
    "Cash conversion is the quality check. Operating cash flow was $102.7 billion against net income of $120.1 billion. Free cash flow, after $6.0 billion of capital expenditure, was $96.7 billion, a 44.8% margin. The gap versus net income is working-capital investment, not accrual earnings. Financing cash flow was −$48.5 billion, driven by $40.1 billion of share repurchases. Earnings are being converted and returned, not capitalized.",
))

story.append(P("4. Financial Analysis", "H1"))
story.append(P(
    "FY2026 revenue was $215.9 billion, up 65.5% from $130.5 billion. Operating margin was 60.4%. Net income was $120.1 billion. Cash and marketable securities were $62.6 billion against $8.5 billion of debt. The current ratio was 3.91x. Trailing-twelve-month revenue through early October 2026 was about $303 billion; the tables below use audited fiscal years so the ratios foot to the 10-K.",
))

story.append(P("4.1 Traditional DuPont", "H2"))
story.append(P("ROE = net margin × asset turnover × equity multiplier, using ending balances so each firm-year multiplies to its reported ROE. Columns are aligned year-ends, not same fiscal labels.", "Formula"))

rows = [[th(x) for x in ["", "YE23", "YE24", "YE25", "YE26"]]]
rows.append([td("Fiscal year used", left=True),
             td("NV FY23 / peers FY22"), td("NV FY24 / peers FY23"),
             td("NV FY25 / peers FY24"), td("NV FY26 / peers FY25")])
for name, book, key in (("NVIDIA", NV, 1), ("AMD", AMD, 2), ("Intel", INTC, 2)):
    rows.append([td(name + " net margin", left=True)] + [td(fmt_pct(book[p[key]]["ni"] / book[p[key]]["rev"])) for p in ALIGNED])
    rows.append([td(name + " asset turnover", left=True)] + [td(fmt_x(book[p[key]]["rev"] / book[p[key]]["ta"])) for p in ALIGNED])
    rows.append([td(name + " equity multiplier", left=True)] + [td(fmt_x(book[p[key]]["ta"] / book[p[key]]["eq"])) for p in ALIGNED])
    rows.append([td(name + " ROE", bold=True, left=True)] + [td(fmt_pct(book[p[key]]["ni"] / book[p[key]]["eq"]), bold=True) for p in ALIGNED])
story.append(table(rows, [1.85*inch, 1.3*inch, 1.3*inch, 1.3*inch, 1.3*inch], highlight_last=True))
story.append(P("Table 1. Traditional DuPont on aligned year-ends. Each ROE equals that row’s margin × turnover × leverage.", "Caption"))
story.append(P(
    "NVIDIA’s YE23 dip (ROE 19.8%) was a margin event, not a leverage event. The recovery is also a margin event: net margin went from 16.2% to 55.6%, while leverage fell from 1.86x to 1.31x. AMD’s ROE stays in single digits after the Xilinx asset step-up. Intel’s is negative in YE25 and about zero in YE26. The YE26 NVIDIA decline from 91.9% to 76.3% is the equity base catching up.",
))
story.append(Image(f"{CHART}/drivers.png", width=6.9*inch, height=2.73*inch))
story.append(P("Figure 1. All three firms, same aligned periods. Margin, not leverage, separates NVIDIA.", "Caption"))

story.append(P("4.2 Alternative DuPont (Palepu–Healy)", "H2"))
story.append(P(
    "ROE = RNOA + (RNOA − NBC) × FLEV, on the same ending equity as Table 1. Net financial expense is the plug that makes after-tax operating profit minus net financial expense equal net income, so the identity is exact. The prior 149.8% RNOA used average balances and did not foot to 76.3%.",
    "Body",
))
story.append(P("NOPAT = operating income × (1 − tax). NFO = debt − cash and securities. NOA = equity + NFO. NBC = NFE / NFO.", "Formula"))

nv_a = advanced(NV[2026]["oi"], NV[2026]["ni"], NV[2026]["eq"], **NV_FIN[2026])
amd_a = advanced(AMD[2025]["oi"], AMD[2025]["ni"], AMD[2025]["eq"], **AMD_FIN[2025])
int_a = advanced(INTC[2025]["oi"], INTC[2025]["ni"], INTC[2025]["eq"], **INTC_FIN[2025])

def adv_cells(a):
    return [
        f"{a['rnoa']*100:.1f}%",
        f"{a['nbc']*100:.1f}%",
        f"{a['spread']*100:.1f}%",
        f"{a['flev']:.2f}x",
        f"{a['effect']*100:.1f}%",
        f"{a['decomp']*100:.1f}%",
        f"{a['roe']*100:.1f}%",
    ]

rows = [[th(x) for x in ["Component", "NVIDIA YE26", "AMD YE26", "Intel YE26"]]]
labels = ["RNOA", "Net borrowing cost, after tax", "Spread (RNOA − NBC)", "Financial leverage (NFO / equity)", "Leverage effect (spread × FLEV)", "ROE, advanced", "ROE, traditional (check)"]
for i, lab in enumerate(labels):
    bold = i >= 5
    vals = [adv_cells(nv_a)[i], adv_cells(amd_a)[i], adv_cells(int_a)[i]]
    rows.append([td(lab, bold=bold, left=True)] + [td(v, bold=bold) for v in vals])
story.append(table(rows, [2.7*inch, 1.4*inch, 1.4*inch, 1.4*inch], highlight_last=True))
story.append(P("Table 2. Advanced DuPont at YE26. The last two rows are the same ROE. NVIDIA 76.3% matches Table 1.", "Caption"))
story.append(P(
    f"NVIDIA’s ending RNOA is {nv_a['rnoa']*100:.1f}%, not 149.8%. Financial leverage is {nv_a['flev']:.2f}x because cash and securities exceed debt. The leverage effect is {nv_a['effect']*100:.1f} points, and {nv_a['rnoa']*100:.1f}% + ({nv_a['effect']*100:.1f}%) = {nv_a['roe']*100:.1f}%, the same 76.3% as net income over ending equity. AMD and Intel foot the same way.",
))
story.append(Image(f"{CHART}/rnoa.png", width=6.9*inch, height=3.02*inch))
story.append(P("Figure 2. RNOA for all three firms on the same aligned year-ends. NVIDIA’s operating return is an order of magnitude above the peers. Intel’s is negative in the last two periods.", "Caption"))
story.append(Image(f"{CHART}/roe.png", width=6.9*inch, height=3.02*inch))
story.append(P("Figure 3. Ending-equity ROE, same alignment. NVIDIA’s YE26 decline is retained earnings. Intel’s trough is YE25.", "Caption"))

story.append(P("5. Prospective Analysis", "H1"))
story.append(P(
    "The valuation is a five-year residual-income model. Value equals current book equity plus the present value of residual income and a terminal value. Residual income is net income minus the equity charge on beginning book equity. No off-balance-sheet asset is added. The prior $1,996 billion “CUDA moat” adjustment is dropped because it was the plug that forced a $280 price.",
))
story.append(P("V0 = B0 + Σ (NIt − re × Bt−1) / (1 + re)^t + terminal value. Terminal value = RI5 × (1 + g) / (re − g).", "Formula"))

story.append(P("5.1 Base forecast", "H2"))
story.append(P(
    "All three firms use the same engine. Five years, a 15% tax rate, net income equal to revenue times the operating margin times 0.85, book equity rolling forward by net income, a 12% equity charge, and a 3.5% terminal growth rate. No plug. Growth and margin paths differ because the businesses differ: NVIDIA fades from 25% growth and a 58% margin, AMD fades from 25% growth as margins rise from 14% to 22%, and Intel grows 2–6% as the margin recovers from 2% to 12%.",
))

fc = FORECASTS
fc_rows = [[th(x) for x in ["Revenue, $ millions", "+1", "+2", "+3", "+4", "+5"]]]
for name in ("NVIDIA", "AMD", "Intel"):
    fc_rows.append([td(name, bold=True, left=True)] + [td(f"{v:,.0f}", bold=True) for v in fc[name]["revs"]])
fc_rows.append([td("Net income", left=True)] + [td("") for _ in range(5)])
for name in ("NVIDIA", "AMD", "Intel"):
    fc_rows.append([td(name, left=True)] + [td(f"{v:,.0f}") for v in fc[name]["nis"]])
story.append(table(fc_rows, [1.7*inch, 1.04*inch, 1.04*inch, 1.04*inch, 1.04*inch, 1.04*inch]))
story.append(P("Table 3. Same five-year forecast for all three. Years are steps after each firm’s latest filing.", "Caption"))
story.append(Image(f"{CHART}/revenue.png", width=6.9*inch, height=3.02*inch))
story.append(P("Figure 4. Solid lines are reported revenue on aligned year-ends. Dashed lines are the common forecast.", "Caption"))

story.append(P("5.2 Value bridge", "H2"))
bridge = [[th(x) for x in ["Component", "NVIDIA", "AMD", "Intel"]]]
def brow(label, fn, bold=False):
    return [td(label, bold=bold, left=True)] + [td(fn(n), bold=bold) for n in ("NVIDIA", "AMD", "Intel")]
bridge += [
    brow("Book equity", lambda n: f"${fc[n]['b0']:,.0f}"),
    brow("PV of residual income", lambda n: f"${fc[n]['pv_ri']:,.0f}"),
    brow("PV of terminal value", lambda n: f"${fc[n]['pv_tv']:,.0f}"),
    brow("Equity value", lambda n: f"${fc[n]['equity']:,.0f}", bold=True),
    brow("Diluted shares (millions)", lambda n: f"{SHARES[n]:,.0f}"),
    brow("Value per share", lambda n: f"${fc[n]['equity']/SHARES[n]:,.0f}", bold=True),
    brow("Price, October 5, 2026", lambda n: f"${PRICES[n]:,.0f}"),
]
story.append(table(bridge, [2.2*inch, 1.55*inch, 1.55*inch, 1.55*inch]))
nv_ps = fc["NVIDIA"]["equity"] / SHARES["NVIDIA"]
amd_ps = fc["AMD"]["equity"] / SHARES["AMD"]
int_ps = fc["Intel"]["equity"] / SHARES["Intel"]
story.append(P(f"Table 4. Same residual-income bridge. Values per share are ${nv_ps:.0f}, ${amd_ps:.0f}, and ${int_ps:.0f}.", "Caption"))

story.append(P("5.3 What would have to be true", "H2"))
story.append(P(
    "Table 5 reprices the same forecasts at other costs of equity. Terminal growth stays at 3.5%. No firm clears its October 5 price on this filing-based engine.",
))
sc = [[th(x) for x in ["Cost of equity", "NVIDIA", "AMD", "Intel"]]]
for re_s in (0.105, 0.12, 0.135):
    cells = [td(f"{re_s*100:.1f}%" + (" (base)" if abs(re_s - 0.12) < 1e-9 else ""), left=True, bold=abs(re_s - 0.12) < 1e-9)]
    specs = {
        "NVIDIA": (215938, 157293, [0.25, 0.20, 0.14, 0.10, 0.07], [0.58, 0.55, 0.52, 0.50, 0.48]),
        "AMD": (34639, 62999, [0.25, 0.20, 0.16, 0.13, 0.10], [0.14, 0.17, 0.19, 0.21, 0.22]),
        "Intel": (52853, 126360, [0.02, 0.04, 0.06, 0.05, 0.04], [0.02, 0.06, 0.09, 0.11, 0.12]),
    }
    for name in ("NVIDIA", "AMD", "Intel"):
        rev0, bv0, growth, margins = specs[name]
        alt = run_forecast(rev0, bv0, growth, margins, re=re_s)
        cells.append(td(f"${alt['equity']/SHARES[name]:,.0f}", bold=abs(re_s - 0.12) < 1e-9))
    sc.append(cells)
story.append(table(sc, [1.7*inch, 1.7*inch, 1.7*inch, 1.7*inch]))
story.append(P("Table 5. Value per share under the same forecast, varying only the cost of equity.", "Caption"))
story.append(P(
    f"Hold on NVIDIA. The common model puts it at ${nv_ps:.0f} against $236, AMD at ${amd_ps:.0f} against $632, and Intel at ${int_ps:.0f} against $117. The $280 target is not in this table. Peer prices are farther above the same engine, which does not make NVIDIA cheap.",
))

story.append(Spacer(1, 6))
story.append(P(
    "Sources: NVIDIA, AMD, and Intel annual reports in the project filing set; curated statement extract and DuPont, forecast, and scenario files in nathanalam/bav-nvidia. Market price and trailing revenue as of October 5, 2026. Dollar figures in millions except per-share amounts.",
    "Note",
))

doc = SimpleDocTemplate(
    OUT, pagesize=letter,
    leftMargin=0.7*inch, rightMargin=0.7*inch,
    topMargin=0.62*inch, bottomMargin=0.62*inch,
    title="NVIDIA Accounting and Performance Analysis",
    author="Nathan Shiham Alam",
)
doc.build(story, onFirstPage=header_footer, onLaterPages=header_footer)
print("wrote", OUT)
