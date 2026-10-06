"""Quarterly DuPont, trailing-twelve-month comparison, and twenty-quarter forecast.

The identity is computed on the quarter. Annualized figures are the quarter times
four, so a return can sit next to an annual cost of equity. Trailing twelve months
uses the sum of four quarters over the average of the opening and closing balance.
"""

from __future__ import annotations

from quarterly_statements import ORDER, QUARTERS, SHARES, SLOTS, TAX


def net_debt(row: dict) -> float:
    return row["debt"] + row["lease"] - row["cash"] - row["liq"]


def decompose(prev: dict, cur: dict) -> dict:
    avg_ta = (prev["ta"] + cur["ta"]) / 2
    avg_eq = (prev["eq"] + cur["eq"]) / 2
    nd0, nd1 = net_debt(prev), net_debt(cur)
    avg_nd = (nd0 + nd1) / 2
    avg_noa = avg_eq + avg_nd
    npm = cur["ni"] / cur["rev"]
    ato = cur["rev"] / avg_ta
    em = avg_ta / avg_eq
    roe = cur["ni"] / avg_eq
    if abs(npm * ato * em - roe) > 1e-9:
        raise AssertionError((cur["end"], npm * ato * em, roe))
    alt = cur["oi"] * (1 - TAX) / avg_noa
    out = dict(
        npm=npm, ato=ato, em=em, roe=roe, opm=cur["oi"] / cur["rev"],
        capex_rev=cur["capex"] / cur["rev"],
        ocf_capex=(cur["ocf"] / cur["capex"] if cur["capex"] else None),
        nd=nd1, avg_nd=avg_nd, avg_noa=avg_noa, avg_eq=avg_eq, avg_ta=avg_ta,
        nd_eq=nd1 / cur["eq"], roe_ann=roe * 4, ato_ann=ato * 4,
        alt=alt, alt_ann=alt * 4,
        pm=None, turn=None, rnoa=None, rnoa_ann=None, nbc=None,
        spread=None, flev=None, gain=None, nopat=None, nfe=None,
    )
    if cur["ii"] is None:
        return out
    nfe = (cur["ix"] - cur["ii"]) * (1 - TAX)
    nopat = cur["ni"] + nfe
    pm = nopat / cur["rev"]
    turn = cur["rev"] / avg_noa
    rnoa = nopat / avg_noa
    if abs(pm * turn - rnoa) > 1e-9:
        raise AssertionError(cur["end"])
    flev = avg_nd / avg_eq
    near = abs(avg_nd) < 250
    nbc = None if near else nfe / avg_nd
    if near:
        spread = gain = None
    else:
        spread = rnoa - nbc
        gain = spread * flev
        if abs((rnoa + gain) - roe) > 1e-6:
            raise AssertionError((cur["end"], rnoa + gain, roe))
    out.update(
        pm=pm, turn=turn, rnoa=rnoa, rnoa_ann=rnoa * 4, nbc=nbc,
        spread=spread, flev=flev, gain=gain, nopat=nopat, nfe=nfe,
    )
    return out


def ratio_rows(name: str) -> list[dict]:
    rows = QUARTERS[name]
    return rows[1:]


def ratios_for(name: str) -> list[dict]:
    rows = QUARTERS[name]
    return [decompose(rows[i], rows[i + 1]) for i in range(len(rows) - 1)]


RATIOS = {name: ratios_for(name) for name in ORDER}


def ttm_at(name: str, slot_index: int) -> dict | None:
    """slot_index is the position in SLOTS. Needs four ratio quarters ending there."""
    if slot_index < 3:
        return None
    rows = QUARTERS[name]
    # rows[0] is the opening quarter. Ratio quarter i is rows[i + 1].
    end_row = slot_index + 1
    start_balance = rows[end_row - 4]
    window = rows[end_row - 3: end_row + 1]
    if len(window) != 4:
        raise AssertionError(name)
    rev = sum(row["rev"] for row in window)
    oi = sum(row["oi"] for row in window)
    ni = sum(row["ni"] for row in window)
    ocf = sum(row["ocf"] for row in window)
    capex = sum(row["capex"] for row in window)
    end = rows[end_row]
    avg_eq = (start_balance["eq"] + end["eq"]) / 2
    avg_ta = (start_balance["ta"] + end["ta"]) / 2
    nd0, nd1 = net_debt(start_balance), net_debt(end)
    avg_nd = (nd0 + nd1) / 2
    avg_noa = avg_eq + avg_nd
    interest = None if any(row["ii"] is None for row in window) else (
        sum(row["ix"] - row["ii"] for row in window) * (1 - TAX)
    )
    nopat = None if interest is None else ni + interest
    return dict(
        slot=SLOTS[slot_index],
        rev=rev, oi=oi, ni=ni, ocf=ocf, capex=capex,
        opm=oi / rev, npm=ni / rev, roe=ni / avg_eq,
        ato=rev / avg_ta, em=avg_ta / avg_eq,
        rnoa=None if nopat is None else nopat / avg_noa,
        alt=oi * (1 - TAX) / avg_noa,
        nd=nd1, avg_eq=avg_eq,
    )


def latest_ttm(name: str) -> dict:
    return ttm_at(name, len(SLOTS) - 1)


# Twenty explicit quarters, five years. NVIDIA's first quarter is the company's
# guide. Every later quarter steps off the quarter before it. Operating margin
# starts at the trailing-twelve-month margin and glides to the target.
# Net income is operating income times a clean conversion: NVIDIA's latest
# quarter, and 1 - 21% for AMD and Intel. Equity-security gains and Intel's
# escrowed-share mark are not forecast to recur. Cost of equity and terminal
# growth are the rates already used on the annual model.
def _steps(groups: list[tuple[int, float]]) -> list[float]:
    rates: list[float] = []
    for count, rate in groups:
        rates.extend([rate] * count)
    return rates


FORECAST_ASSUMPTIONS = {
    "NVIDIA": dict(
        ke=0.135, g=0.03, om_target=0.55, ni_on_oi=59688 / 63734,
        growth=[None, *_steps([
            (1, 0.06), (1, 0.05), (1, 0.04), (4, 0.035), (4, 0.025), (4, 0.018), (4, 0.012),
        ])],
        guided_revenue=108_000,
        note="First quarter is the $108bn guide, midpoint, with no China data-center compute. Margin glides from the trailing-twelve-month 65% toward 55%.",
    ),
    "AMD": dict(
        ke=0.125, g=0.03, om_target=0.22, ni_on_oi=1 - TAX,
        growth=_steps([
            (1, 0.06), (1, 0.05), (1, 0.045), (1, 0.04), (4, 0.035), (4, 0.028), (4, 0.02), (4, 0.015),
        ]),
        guided_revenue=None,
        note="Sequential growth steps down from 6%. Net income is operating income taxed at 21%, because quarterly interest income is not disclosed apart from other income.",
    ),
    "Intel": dict(
        ke=0.097, g=0.03, om_target=0.10, ni_on_oi=1 - TAX,
        growth=_steps([(4, 0.02), (4, 0.015), (4, 0.012), (4, 0.01), (4, 0.008)]),
        guided_revenue=None,
        note="Margin starts at the trailing-twelve-month result, about breakeven, and glides to 10%. The escrowed-share mark is not repeated.",
    ),
}


def forecast(name: str, *, om0: float | None = None, om_target: float | None = None) -> dict:
    assumptions = FORECAST_ASSUMPTIONS[name]
    if len(assumptions["growth"]) != 20:
        raise AssertionError((name, len(assumptions["growth"])))
    history = QUARTERS[name]
    last = history[-1]
    ttm = latest_ttm(name)
    ke = assumptions["ke"]
    ke_q = ke / 4
    g_q = (1 + assumptions["g"]) ** 0.25 - 1
    if ke_q <= g_q:
        raise AssertionError(name)
    if om0 is None:
        om0 = ttm["opm"]
    om1 = assumptions["om_target"] if om_target is None else om_target
    ato = (last["rev"] * 4) / last["ta"]
    eq_ratio = last["eq"] / last["ta"]
    revenue = []
    previous = last["rev"]
    for step, growth in enumerate(assumptions["growth"]):
        if step == 0 and assumptions["guided_revenue"] is not None:
            nxt = assumptions["guided_revenue"]
        else:
            nxt = previous * (1 + growth)
        revenue.append(nxt)
        previous = nxt
    quarters = []
    equity_prev = last["eq"]
    pv_ri = 0.0
    for t, rev in enumerate(revenue, start=1):
        om = om0 + (om1 - om0) * t / len(revenue)
        oi = rev * om
        ni = oi * assumptions["ni_on_oi"]
        ta = (rev * 4) / ato
        eq = ta * eq_ratio
        residual = ni - ke_q * equity_prev
        discount = (1 + ke_q) ** t
        pv_ri += residual / discount
        quarters.append(dict(
            t=t, rev=rev, om=om, oi=oi, ni=ni, ta=ta, eq=eq,
            residual=residual, roe=ni / equity_prev,
        ))
        equity_prev = eq
    last_ri = quarters[-1]["residual"]
    terminal = last_ri * (1 + g_q) / (ke_q - g_q)
    pv_tv = terminal / (1 + ke_q) ** len(quarters)
    mve = last["eq"] + pv_ri + pv_tv
    shares = SHARES[name]["diluted_m"]
    return dict(
        assumptions=assumptions, quarters=quarters, mve=mve, price=mve / shares,
        shares=shares, pv_ri=pv_ri, pv_tv=pv_tv, book=last["eq"],
        ke=ke, g=assumptions["g"], om0=om0, ato=ato, eq_ratio=eq_ratio,
    )


FORECASTS = {name: forecast(name) for name in ORDER}


def year_totals(name: str) -> list[dict]:
    """Five forecast years, four quarters each."""
    rows = []
    for k in range(5):
        chunk = FORECASTS[name]["quarters"][k * 4:(k + 1) * 4]
        rev = sum(q["rev"] for q in chunk)
        oi = sum(q["oi"] for q in chunk)
        ni = sum(q["ni"] for q in chunk)
        residual = sum(q["residual"] for q in chunk)
        rows.append(dict(rev=rev, oi=oi, ni=ni, residual=residual, om=oi / rev, om_end=chunk[-1]["om"]))
    return rows


def _check() -> None:
    nv = latest_ttm("NVIDIA")
    if abs(nv["rev"] - 302969) > 1 or abs(nv["roe"] - 1.172) > 0.002:
        raise AssertionError((nv["rev"], nv["roe"]))
    for name in ORDER:
        for row in RATIOS[name]:
            if row["rnoa"] is not None and abs((row["rnoa"] + row["gain"]) - row["roe"]) > 1e-6:
                raise AssertionError(name)
    amd = RATIOS["AMD"][-1]
    if amd["rnoa"] is not None:
        raise AssertionError("AMD modified DuPont should be open")
    for name in ORDER:
        if len(FORECASTS[name]["quarters"]) != 20:
            raise AssertionError(name)
        if len(year_totals(name)) != 5:
            raise AssertionError(name)


_check()
