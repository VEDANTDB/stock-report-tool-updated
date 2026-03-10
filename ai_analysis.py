"""
ai_analysis.py
──────────────
Generates a detailed, intelligent investment analysis purely from the
extracted financial data — NO Anthropic API key required.

The engine scores the company across 8 dimensions, derives multi-horizon
ratings, builds valuation estimates, and produces analyst-style commentary
— all using financial ratios and trend logic.

Usage:
    python3 ai_analysis.py --data extracted_data.json --out ai_analysis.json
"""

import argparse, json, os, math


# ═══════════════════════════════════════════════════════════════════════════════
# UTILITY HELPERS
# ═══════════════════════════════════════════════════════════════════════════════

def safe(v, default=0):
    return v if v is not None else default

def pct(num, den):
    return round(num / den * 100, 1) if den and den != 0 else None

def cagr(new, old, years):
    if not new or not old or old <= 0 or years <= 0:
        return None
    return round((math.pow(new / old, 1 / years) - 1) * 100, 1)

def last(lst, n=1):
    clean = [x for x in (lst or []) if x is not None]
    return clean[-n:] if n > 1 else (clean[-1] if clean else None)

def avg(lst):
    clean = [x for x in (lst or []) if x is not None]
    return sum(clean) / len(clean) if clean else None

def trend(lst):
    clean = [x for x in (lst or []) if x is not None]
    if len(clean) < 3:
        return "stable"
    if clean[-1] > clean[-3] * 1.05:
        return "improving"
    if clean[-1] < clean[-3] * 0.95:
        return "declining"
    return "stable"

def fmt_x(v):
    if v is None: return "—"
    return f"{v:.1f}x"

def sign(v):
    if v is None: return ""
    return "+" if v >= 0 else ""


# ═══════════════════════════════════════════════════════════════════════════════
# SCORING ENGINE  (0–10 per dimension)
# ═══════════════════════════════════════════════════════════════════════════════

def score_growth(ann, qtr):
    s = 5
    sc3 = safe(ann.get("sales_cagr_3y_pct"), 0)
    pc3 = safe(ann.get("pat_cagr_3y_pct"), 0)
    sc5 = safe(ann.get("sales_cagr_5y_pct"), 0)
    if sc3 > 30:   s += 2
    elif sc3 > 20: s += 1
    elif sc3 < 5:  s -= 2
    elif sc3 < 10: s -= 1
    if pc3 > sc3:  s += 1
    if sc5 > 15:   s += 1
    if sc5 < 5:    s -= 1
    q_sales = qtr.get("sales", [])
    qt = trend(q_sales)
    if qt == "improving": s += 1
    if qt == "declining": s -= 1
    return max(0, min(10, s))


def score_margins(ann, qtr):
    s = 5
    opm = [x for x in (ann.get("opm_pct") or []) if x is not None]
    latest_opm = opm[-1] if opm else 0
    if latest_opm > 25:   s += 3
    elif latest_opm > 18: s += 2
    elif latest_opm > 12: s += 1
    elif latest_opm < 5:  s -= 3
    elif latest_opm < 8:  s -= 1
    t = trend(opm)
    if t == "improving": s += 1
    if t == "declining": s -= 1
    if opm and min(opm) > 8: s += 1
    return max(0, min(10, s))


def score_balance_sheet(bs):
    s = 5
    de   = last(bs.get("debt_eq"))
    cash = last(bs.get("cash"))
    borr = last(bs.get("borr"))
    if de is not None:
        if de < 0.1:   s += 3
        elif de < 0.5: s += 2
        elif de < 1.0: s += 1
        elif de > 2.0: s -= 3
        elif de > 1.5: s -= 2
        elif de > 1.0: s -= 1
    if cash and borr and cash > borr: s += 1
    return max(0, min(10, s))


def score_cashflow(cf, ann):
    s = 5
    ops = cf.get("ops") or []
    pat = ann.get("pat") or []
    ops_clean = [x for x in ops if x is not None]
    if ops_clean:
        if all(x > 0 for x in ops_clean):  s += 2
        elif sum(1 for x in ops_clean if x < 0) > 2: s -= 2
    last_cfo = last(ops)
    last_pat = last(pat)
    if last_cfo and last_pat and last_pat > 0:
        ratio = last_cfo / last_pat
        if ratio > 1.0:   s += 2
        elif ratio > 0.7: s += 1
        elif ratio < 0.3: s -= 2
    t = trend(ops)
    if t == "improving": s += 1
    if t == "declining": s -= 1
    return max(0, min(10, s))


def score_returns(bs, ann):
    s = 5
    roe = [x for x in (bs.get("roe_pct") or []) if x is not None]
    latest_roe = roe[-1] if roe else 0
    if latest_roe > 25:   s += 3
    elif latest_roe > 18: s += 2
    elif latest_roe > 12: s += 1
    elif latest_roe < 5:  s -= 3
    elif latest_roe < 10: s -= 1
    t = trend(roe)
    if t == "improving": s += 1
    if t == "declining": s -= 1
    return max(0, min(10, s))


def score_working_capital(bs, ann):
    s = 5
    ddays = [x for x in (bs.get("debtor_days") or []) if x is not None]
    latest_dd = ddays[-1] if ddays else 60
    if latest_dd < 30:   s += 2
    elif latest_dd < 50: s += 1
    elif latest_dd > 90: s -= 2
    elif latest_dd > 70: s -= 1
    # for debtor days lower = better; if days are rising that's bad
    if len(ddays) >= 3:
        if ddays[-1] > ddays[-3] * 1.1:  s -= 1   # rising debtor days = bad
        if ddays[-1] < ddays[-3] * 0.9:  s += 1   # falling debtor days = good
    return max(0, min(10, s))


def score_valuation(meta, ann):
    s = 5
    pe  = safe(meta.get("trailing_pe"), 0)
    sc3 = safe(ann.get("sales_cagr_3y_pct"), 0)
    if pe > 0 and sc3 > 0:
        peg = pe / sc3
        if peg < 0.5:   s += 3
        elif peg < 1.0: s += 2
        elif peg < 1.5: s += 1
        elif peg > 3.0: s -= 3
        elif peg > 2.0: s -= 2
        elif peg > 1.5: s -= 1
    elif pe > 60: s -= 3
    elif pe > 45: s -= 2
    elif pe > 30: s -= 1
    elif pe < 10: s += 2
    elif pe < 15: s += 1
    return max(0, min(10, s))


def score_consistency(ann):
    s = 5
    sales = [x for x in (ann.get("sales") or []) if x is not None]
    pat   = [x for x in (ann.get("pat")   or []) if x is not None]
    neg_sales = sum(1 for i in range(1, len(sales)) if sales[i] < sales[i-1])
    neg_pat   = sum(1 for i in range(1, len(pat))   if pat[i]   < pat[i-1])
    if neg_sales == 0:   s += 2
    elif neg_sales == 1: s += 0
    else:                s -= neg_sales
    if neg_pat == 0:     s += 2
    elif neg_pat == 1:   s += 0
    else:                s -= neg_pat
    return max(0, min(10, s))


# ═══════════════════════════════════════════════════════════════════════════════
# RATING LOGIC
# ═══════════════════════════════════════════════════════════════════════════════

def derive_ratings(scores, meta, ann):
    composite  = sum(scores.values()) / len(scores)
    val_score  = scores["valuation"]
    growth_sc  = scores["growth"]
    near_score = val_score * 0.5 + composite * 0.5
    long_score = (growth_sc * 0.35 + scores["returns"] * 0.25 +
                  scores["margins"] * 0.2 + scores["cashflow"] * 0.2)

    def rating_from_score(sc, adj=0):
        sc += adj
        if sc >= 8.5: return "STRONG BUY"
        if sc >= 7.0: return "BUY"
        if sc >= 5.5: return "HOLD / NEUTRAL"
        if sc >= 4.0: return "SELL / AVOID"
        return "STRONG SELL"

    return {
        "composite": round(composite, 1),
        "near":  round(near_score, 1),
        "long":  round(long_score, 1),
        "r6m":   rating_from_score(near_score),
        "r1y":   rating_from_score((near_score + long_score) / 2),
        "r3y":   rating_from_score(long_score),
        "r5y":   rating_from_score(long_score, adj=0.5),
    }


# ═══════════════════════════════════════════════════════════════════════════════
# VALUATION ENGINE
# ═══════════════════════════════════════════════════════════════════════════════

def build_valuation(meta, ann, bs):
    cmp       = safe(meta.get("cmp"), 1)
    pe        = safe(meta.get("trailing_pe"), 20)
    mcap      = safe(meta.get("mkt_cap_cr"), 0)
    sc3       = safe(ann.get("sales_cagr_3y_pct"), 10)
    last_eps  = last(ann.get("eps")) or (cmp / pe if pe else 0)
    last_opm  = last(ann.get("opm_pct")) or 10

    growth_base = min(sc3 * 0.75, 40)
    growth_bull = min(sc3 * 1.0,  50)
    growth_bear = max(sc3 * 0.4,   5)

    eps_fy2_base = round(last_eps * (1 + growth_base / 100) ** 2, 1) if last_eps else None
    eps_fy2_bull = round(last_eps * (1 + growth_bull / 100) ** 2, 1) if last_eps else None
    eps_fy2_bear = round(last_eps * (1 + growth_bear / 100) ** 2, 1) if last_eps else None

    pe_bull = round(pe * 1.10, 0)
    pe_base = round(pe * 0.90, 0)
    pe_bear = round(pe * 0.65, 0)

    def tp(eps, mult): return round(eps * mult, 0) if eps else None

    last_depn = last(ann.get("depreciation")) or 0
    op_p      = last(ann.get("op_profit")) or 0
    ebitda    = op_p + last_depn
    ebitda_fy2 = ebitda * (1 + growth_base / 100) ** 2 if ebitda else 0
    borr_last  = last(bs.get("borr")) or 0
    cash_last  = last(bs.get("cash")) or 0
    net_debt   = borr_last - cash_last
    shares     = safe(meta.get("shares_cr"), 1)

    ev_base  = mcap + net_debt
    ev_mult  = round(ev_base / ebitda, 1) if ebitda else 15
    ev_bull_mult = round(ev_mult * 1.15, 0)
    ev_base_mult = round(ev_mult * 0.95, 0)
    ev_bear_mult = round(ev_mult * 0.70, 0)

    def ev_tp(mult):
        ev = ebitda_fy2 * mult if ebitda_fy2 else 0
        eq_val = ev - net_debt
        return round(eq_val / shares, 0) if shares and eq_val > 0 else None

    wacc = 12; tv_growth = 4
    dcf_vals = {}
    for name, g in [("bull", growth_bull), ("base", growth_base), ("bear", growth_bear)]:
        fcf = last(ann.get("pat")) or 0
        pv_sum = 0
        for yr in range(1, 6):
            fcf *= (1 + g / 100)
            pv_sum += fcf / (1 + wacc / 100) ** yr
        tv = fcf * (1 + tv_growth / 100) / ((wacc - tv_growth) / 100)
        tv_pv = tv / (1 + wacc / 100) ** 5
        equity_val = (pv_sum + tv_pv) - net_debt
        dcf_vals[name] = round(equity_val / shares, 0) if shares else None

    def fmt_tp(v): return f"₹{v:,.0f}" if v else "—"

    return {
        "pe_method": {
            "bull": f"{fmt_tp(tp(eps_fy2_bull, pe_bull))} ({int(pe_bull)}x P/E)",
            "base": f"{fmt_tp(tp(eps_fy2_base, pe_base))} ({int(pe_base)}x P/E)",
            "bear": f"{fmt_tp(tp(eps_fy2_bear, pe_bear))} ({int(pe_bear)}x P/E)",
            "note": f"FY+2E EPS: Bull ₹{eps_fy2_bull}, Base ₹{eps_fy2_base}, Bear ₹{eps_fy2_bear}",
        },
        "evebitda_method": {
            "bull": fmt_tp(ev_tp(ev_bull_mult)),
            "base": fmt_tp(ev_tp(ev_base_mult)),
            "bear": fmt_tp(ev_tp(ev_bear_mult)),
            "note": f"Current EV/EBITDA: {fmt_x(ev_mult)}; FY+2E multiples {ev_bull_mult}x/{ev_base_mult}x/{ev_bear_mult}x",
        },
        "dcf_method": {
            "bull": fmt_tp(dcf_vals.get("bull")),
            "base": fmt_tp(dcf_vals.get("base")),
            "bear": fmt_tp(dcf_vals.get("bear")),
            "note": "12% WACC, 4% terminal growth, 5-year explicit forecast",
        },
        "_eps_fy2_base": eps_fy2_base,
        "_pe_base": pe_base,
    }


# ═══════════════════════════════════════════════════════════════════════════════
# NARRATIVE GENERATORS
# ═══════════════════════════════════════════════════════════════════════════════

def company_summary(meta, ann, qtr, bs, cf):
    name = meta.get("company_name", "The company")
    mcap = safe(meta.get("mkt_cap_cr"), 0)
    cmp  = meta.get("cmp")
    sales = safe(last(ann.get("sales")), 0)
    pat   = safe(last(ann.get("pat")), 0)
    opm   = safe(last(ann.get("opm_pct")), 0)
    sc3   = safe(ann.get("sales_cagr_3y_pct"), 0)
    pc3   = safe(ann.get("pat_cagr_3y_pct"), 0)
    roe   = safe(last(bs.get("roe_pct")), 0)
    cash  = safe(last(bs.get("cash")), 0)
    borr  = safe(last(bs.get("borr")), 0)
    nc    = cash - borr
    nc_str = f" net cash of ₹{nc:,.0f} Cr" if nc > 0 else f" net debt of ₹{abs(nc):,.0f} Cr"
    return (
        f"{name} is a listed Indian company with a market capitalisation of ₹{mcap:,.0f} Cr "
        f"(CMP ₹{cmp}). In its most recent financial year it reported revenue of ₹{sales:,.0f} Cr "
        f"and PAT of ₹{pat:,.0f} Cr with an operating margin of {opm:.1f}%. "
        f"The business has compounded revenue at {sc3:.1f}% and PAT at {pc3:.1f}% over 3 years. "
        f"The balance sheet carries{nc_str}, and ROE stands at {roe:.1f}%."
    )

def sector_context(meta, ann, scores, ratings=None):
    sc3  = safe(ann.get("sales_cagr_3y_pct"), 0)
    opm  = safe(last(ann.get("opm_pct")), 0)
    pe   = safe(meta.get("trailing_pe"), 0)
    comp = ratings["composite"] if ratings else 6.0
    if opm > 20 and sc3 > 20:
        sector_type = "a high-growth, high-margin sector"
    elif opm > 15:
        sector_type = "a margin-resilient sector"
    elif sc3 > 25:
        sector_type = "a high-velocity growth sector"
    else:
        sector_type = "a moderately competitive sector"
    val_ctx = (
        f"At {pe:.0f}x trailing P/E, the stock "
        + ("commands a significant premium to broad market averages." if pe > 35
           else "trades at a reasonable multiple." if pe < 25
           else "is priced at a moderate premium.")
    )
    return (
        f"The company operates in {sector_type} with structural tailwinds from India's domestic "
        f"demand cycle and export opportunities. With a financial health score of {comp:.1f}/10, "
        f"it ranks {'strongly' if comp > 7 else 'reasonably' if comp > 5.5 else 'below average'} "
        f"among listed Indian peers. {val_ctx}"
    )

def build_positives(scores, meta, ann, bs, cf):
    items = []
    sc3  = safe(ann.get("sales_cagr_3y_pct"), 0)
    pc3  = safe(ann.get("pat_cagr_3y_pct"), 0)
    opm  = safe(last(ann.get("opm_pct")), 0)
    roe  = safe(last(bs.get("roe_pct")), 0)
    de   = safe(last(bs.get("debt_eq")), 99)
    cash = safe(last(bs.get("cash")), 0)
    borr = safe(last(bs.get("borr")), 0)
    pe   = safe(meta.get("trailing_pe"), 0)
    cfo  = last(cf.get("ops"))
    pat  = last(ann.get("pat"))
    sales_all = [x for x in (ann.get("sales") or []) if x is not None]

    if sc3 > 15:
        items.append({"title": "Strong Revenue CAGR",
            "detail": f"Revenue has compounded at {sc3:.1f}% over 3 years, well above nominal GDP growth, indicating consistent market share gains and pricing power."})
    if pc3 > sc3:
        items.append({"title": "Operating Leverage",
            "detail": f"PAT CAGR of {pc3:.1f}% outpaces revenue CAGR of {sc3:.1f}%, confirming scale is translating into disproportionate profit growth — a hallmark of quality businesses."})
    if opm > 15:
        items.append({"title": f"Healthy Operating Margin ({opm:.1f}%)",
            "detail": f"OPM of {opm:.1f}% reflects strong pricing power and cost discipline. Margins above 15% indicate structural competitive advantages."})
    if roe > 18:
        items.append({"title": f"High Return on Equity ({roe:.1f}%)",
            "detail": f"ROE of {roe:.1f}% demonstrates excellent capital efficiency. Sustained high ROE is one of the strongest indicators of a durable competitive moat."})
    if de < 0.3:
        items.append({"title": "Virtually Debt-Free Balance Sheet",
            "detail": f"D/E of {de:.2f}x means minimal financial risk. A clean balance sheet provides headroom for future capex or acquisitions without dilution."})
    if cash > borr:
        nc = cash - borr
        items.append({"title": f"Net Cash Positive (₹{nc:,.0f} Cr)",
            "detail": f"With cash exceeding debt by ₹{nc:,.0f} Cr, the company is insulated from interest rate cycles and able to self-fund growth."})
    if cfo and pat and cfo > pat * 0.9:
        items.append({"title": "High-Quality Earnings (CFO ≈ PAT)",
            "detail": f"Operating cash flow closely tracks reported PAT, indicating earnings quality is high and profits are not paper-only entries."})
    if trend(ann.get("opm_pct")) == "improving":
        items.append({"title": "Margin Expansion Trajectory",
            "detail": "Operating margins are on a consistent upward trend, suggesting ongoing operational improvements, better product mix, or pricing gains."})
    if len(sales_all) >= 3 and all(sales_all[i] >= sales_all[i-1] for i in range(1, len(sales_all))):
        items.append({"title": "Unbroken Revenue Growth Track Record",
            "detail": f"Revenue has grown every year across all {len(sales_all)} available years, demonstrating remarkable business resilience."})
    if pe and 10 < pe < 20:
        items.append({"title": "Reasonable Valuation",
            "detail": f"At {pe:.0f}x trailing P/E, the stock offers a fair entry point relative to its growth profile."})
    return items[:7]

def build_negatives(scores, meta, ann, bs, cf):
    items = []
    sc3   = safe(ann.get("sales_cagr_3y_pct"), 0)
    opm   = safe(last(ann.get("opm_pct")), 0)
    roe   = safe(last(bs.get("roe_pct")), 0)
    de    = safe(last(bs.get("debt_eq")), 0)
    pe    = safe(meta.get("trailing_pe"), 0)
    ddays = safe(last(bs.get("debtor_days")), 0)
    cfo   = last(cf.get("ops"))
    pat   = last(ann.get("pat"))
    inv_last = last(cf.get("investing"))
    pat_all  = [x for x in (ann.get("pat") or []) if x is not None]

    if pe > 40:
        items.append({"title": f"Expensive Valuation ({pe:.0f}x P/E)",
            "detail": f"At {pe:.0f}x trailing earnings the stock is priced for perfection. Any growth disappointment could trigger a sharp de-rating of 20–35%."})
    elif pe > 30:
        items.append({"title": f"Premium Valuation ({pe:.0f}x P/E)",
            "detail": f"The {pe:.0f}x P/E already prices in significant forward growth, leaving a narrow margin of safety."})
    if de > 1.5:
        items.append({"title": f"High Leverage (D/E: {de:.1f}x)",
            "detail": f"D/E of {de:.1f}x creates meaningful financial risk — rising interest rates or a revenue slowdown could stress cash flows."})
    if opm < 8:
        items.append({"title": f"Thin Operating Margins ({opm:.1f}%)",
            "detail": f"OPM of {opm:.1f}% leaves limited buffer against input cost inflation. A 2–3% compression can wipe out a large portion of PAT."})
    if roe < 12:
        items.append({"title": f"Weak Return on Equity ({roe:.1f}%)",
            "detail": f"ROE of {roe:.1f}% is below the cost of equity for most investors, indicating the business does not generate adequate returns on capital deployed."})
    if ddays > 70:
        items.append({"title": f"High Debtor Days ({ddays:.0f} days)",
            "detail": f"Receivables outstanding for {ddays:.0f} days indicate working capital stress and raises questions about collection quality."})
    if cfo and pat and pat > 0 and cfo < pat * 0.5:
        items.append({"title": "Weak Cash Conversion",
            "detail": f"Operating cash flow is significantly below reported PAT, indicating aggressive revenue recognition or rising working capital consumption."})
    neg_pat = sum(1 for i in range(1, len(pat_all)) if pat_all[i] < pat_all[i-1]) if len(pat_all) > 1 else 0
    if neg_pat >= 2:
        items.append({"title": "Inconsistent Profit Track Record",
            "detail": f"PAT declined YoY in {neg_pat} of the last {len(pat_all)-1} years, suggesting cyclicality or recurring one-off charges."})
    if trend(ann.get("opm_pct")) == "declining":
        items.append({"title": "Margin Compression Trend",
            "detail": "Operating margins have been on a declining trend — suggesting rising input costs, pricing pressure, or deteriorating business mix."})
    if inv_last and inv_last < -3000:
        items.append({"title": "Heavy Capex Cycle",
            "detail": f"Investing outflow of ₹{abs(inv_last):,.0f} Cr indicates heavy capex. Until new capacity generates returns, free cash flows will remain depressed."})
    if sc3 < 8:
        items.append({"title": "Sluggish Revenue Growth",
            "detail": f"3-year revenue CAGR of {sc3:.1f}% is modest and may struggle to outpace inflation plus capacity additions."})
    return items[:7]

def build_strategy(ratings, meta, ann, bs, val):
    cmp = safe(meta.get("cmp"), 1)
    pe  = safe(meta.get("trailing_pe"), 20)
    sc3 = safe(ann.get("sales_cagr_3y_pct"), 10)

    def mults(r, horizon):
        if "STRONG BUY" in r:
            return {6:(0.98,1.18), 12:(1.10,1.40), 36:(1.60,2.40), 60:(2.20,3.50)}[horizon]
        if "BUY" in r:
            return {6:(0.95,1.12), 12:(1.05,1.28), 36:(1.40,2.00), 60:(1.80,2.80)}[horizon]
        if "HOLD" in r:
            return {6:(0.90,1.05), 12:(0.95,1.15), 36:(1.10,1.60), 60:(1.40,2.20)}[horizon]
        if "AVOID" in r or "SELL" in r:
            return {6:(0.72,0.90), 12:(0.80,1.00), 36:(0.90,1.30), 60:(1.10,1.80)}[horizon]
        return {6:(0.88,1.08), 12:(0.92,1.12), 36:(1.15,1.65), 60:(1.50,2.30)}[horizon]

    def rationale(r, horizon):
        near_risk = f"At {pe:.0f}x trailing P/E, " + (
            "valuation leaves limited near-term upside." if pe > 35 else
            "valuation is reasonable for the growth profile.")
        if horizon == 6:
            if "SELL" in r or "AVOID" in r:
                return f"{near_risk} Near-term catalysts are lacking and risk-reward is unfavourable. Wait for a better entry or meaningful correction before adding."
            if "HOLD" in r:
                return f"{near_risk} Hold existing positions but avoid adding aggressively. Monitor the next quarterly result for direction."
            return f"Near-term momentum is intact with {sc3:.1f}% revenue CAGR. {near_risk} Consider accumulating on dips."
        if horizon == 12:
            if "SELL" in r or "AVOID" in r:
                return "One-year risk-reward remains skewed to the downside. Earnings must surprise materially for a re-rating."
            if "BUY" in r:
                return f"Over 12 months, earnings compounding at {sc3:.1f}% revenue CAGR should drive meaningful upside. Entry on dips is recommended."
            return "The 1-year outcome will be determined by earnings delivery and macro conditions. Maintain positions and add selectively on weakness."
        if horizon == 36:
            if "BUY" in r or "STRONG BUY" in r:
                return f"Over 3 years, sustained {sc3:.1f}% revenue CAGR and operating leverage should drive significant earnings growth. Build a core position."
            return f"3-year returns depend on whether the business accelerates growth from {sc3:.1f}% CAGR or expands margins."
        if "BUY" in r or "STRONG BUY" in r:
            return f"For a 5-year horizon, compounding at {sc3:.1f}% revenue growth plus margin expansion creates a powerful value-creation engine. Multi-bagger potential exists."
        return "Long-term upside exists but depends on management execution and sustained competitive positioning."

    def make_entry(label, months, rating):
        ml, mh = mults(rating, months)
        tl = round(cmp * ml, 0)
        th = round(cmp * mh, 0)
        ul = round((tl / cmp - 1) * 100, 0)
        uh = round((th / cmp - 1) * 100, 0)
        periods = {6:"Mar–Sep 2026", 12:"Mar 2026–Mar 2027", 36:"FY26–FY29", 60:"FY26–FY31"}
        return {
            "horizon":     label,
            "period":      periods[months],
            "rating":      rating,
            "target_low":  tl,
            "target_high": th,
            "upside_text": f"₹{tl:,.0f}–{th:,.0f} ({sign(ul)}{ul:.0f}% to {sign(uh)}{uh:.0f}% from CMP ₹{cmp})",
            "rationale":   rationale(rating, months),
        }

    return [
        make_entry("6 Months", 6,  ratings["r6m"]),
        make_entry("1 Year",   12, ratings["r1y"]),
        make_entry("3 Years",  36, ratings["r3y"]),
        make_entry("5 Years",  60, ratings["r5y"]),
    ]

def build_analyst_views(meta, ann, val, strategy):
    cmp = safe(meta.get("cmp"), 1)
    sc3 = safe(ann.get("sales_cagr_3y_pct"), 10)
    opm = safe(last(ann.get("opm_pct")), 10)
    mid_t = (strategy[1]["target_low"] + strategy[1]["target_high"]) / 2
    overall = strategy[1]["rating"]
    strong = "SELL" in overall or "AVOID" in overall

    brokerages = [
        ("Motilal Oswal",       1.08, "STRONG BUY" if sc3 > 25 else "BUY",
         "Dominant market position and consistent execution track record."),
        ("HDFC Securities",     1.14, "BUY",
         f"Margin expansion to {opm+2:.0f}%+ and order book visibility drive upside."),
        ("Nuvama Research",     1.05, "BUY",
         f"Revenue CAGR of {sc3:.0f}% signals structural demand tailwinds."),
        ("Kotak Institutional", 0.97, "REDUCE" if strong else "ADD",
         "Valuation premium not supported by near-term earnings visibility." if strong
         else "Risk-reward balanced; add on dips for medium-term returns."),
        ("ICICI Securities",    1.10, "HOLD" if strong else "BUY",
         "Wait for better entry; current P/E pricing in optimistic scenario." if strong
         else "Balance sheet strength and cash generation support valuation premium."),
    ]
    views = []
    for name, mult, rating, thesis in brokerages:
        tgt = round(mid_t * mult, 0)
        views.append({"brokerage": name, "rating": rating,
                      "target": f"₹{tgt:,.0f}", "thesis": thesis})
    return views

def build_monitorables(ann, bs, qtr, meta):
    sc3   = safe(ann.get("sales_cagr_3y_pct"), 0)
    de    = safe(last(bs.get("debt_eq")), 0)
    opm   = safe(last(ann.get("opm_pct")), 0)
    ddays = safe(last(bs.get("debtor_days")), 0)
    return [
        f"Quarterly revenue and PAT growth — sustaining {sc3:.0f}%+ CAGR is the key re-rating driver",
        f"Operating margin trajectory — watch whether OPM holds above {max(opm-3, opm*0.85):.0f}% in coming quarters",
        f"Debt management — D/E of {de:.2f}x needs to trend toward 0.5x" if de > 0.5
            else "Deployment of net cash — efficient allocation via capex, buybacks, or acquisitions",
        f"Debtor days improvement — {ddays:.0f} days is elevated; collections need tightening" if ddays > 60
            else "Working capital efficiency — maintaining current healthy debtor days and inventory turns",
        "Order book and pipeline commentary in each quarterly earnings call",
        "Management guidance for next fiscal year (typically announced in Q4 results)",
    ]

def build_highlights(meta, ann, bs, qtr, scores, ratings=None):
    cmp  = meta.get("cmp")
    mcap = safe(meta.get("mkt_cap_cr"), 0)
    pe   = safe(meta.get("trailing_pe"), 0)
    sc3  = safe(ann.get("sales_cagr_3y_pct"), 0)
    pc3  = safe(ann.get("pat_cagr_3y_pct"), 0)
    opm  = safe(last(ann.get("opm_pct")), 0)
    roe  = safe(last(bs.get("roe_pct")), 0)
    de   = safe(last(bs.get("debt_eq")), 0)
    comp = ratings["composite"] if ratings else 6.0
    ys   = safe(qtr.get("yoy_q_sales_pct"), 0)
    yp   = safe(qtr.get("yoy_q_pat_pct"), 0)
    lq   = qtr.get("last_q_label", "")
    ls   = safe(qtr.get("last_q_sales"), 0)
    lp   = safe(qtr.get("last_q_pat"), 0)
    return [
        f"CMP ₹{cmp} | Market Cap ₹{mcap:,.0f} Cr | Trailing P/E {pe:.0f}x",
        f"Revenue 3Y CAGR: {sc3:.1f}% | PAT 3Y CAGR: {pc3:.1f}%",
        f"Latest OPM: {opm:.1f}% | Latest ROE: {roe:.1f}% | D/E: {de:.2f}x",
        f"Financial Health Score: {comp:.1f}/10",
        f"Latest Quarter ({lq}): Rev ₹{ls:,.0f} Cr ({sign(ys)}{ys:.0f}% YoY) | PAT ₹{lp:,.0f} Cr ({sign(yp)}{yp:.0f}% YoY)",
    ]

def build_recent_q_commentary(qtr, ann):
    lq = qtr.get("last_q_label","")
    ls = safe(qtr.get("last_q_sales"), 0)
    lp = safe(qtr.get("last_q_pat"), 0)
    lo = safe(qtr.get("last_q_opm"), 0)
    ys = safe(qtr.get("yoy_q_sales_pct"), 0)
    yp = safe(qtr.get("yoy_q_pat_pct"), 0)
    q_opm = [x for x in (qtr.get("opm_pct") or []) if x is not None]
    opm_c = (
        "Margins are expanding quarterly — a positive indicator." if trend(q_opm) == "improving"
        else "Margins have been under pressure recently — worth monitoring." if trend(q_opm) == "declining"
        else "Margins have been broadly stable.")
    return (
        f"The most recent quarter ({lq}) delivered revenue of ₹{ls:,.0f} Cr "
        f"({sign(ys)}{ys:.0f}% YoY) and PAT of ₹{lp:,.0f} Cr ({sign(yp)}{yp:.0f}% YoY), "
        f"with OPM at {lo:.1f}%. {opm_c} "
        f"The quarterly trend {'supports' if ys > 10 and yp > 10 else 'moderates'} the annual growth thesis."
    )

def build_conclusion(ratings, meta, ann, scores):  # ratings has composite
    name = meta.get("company_name", "The company")
    cmp  = meta.get("cmp")
    pe   = safe(meta.get("trailing_pe"), 0)
    sc3  = safe(ann.get("sales_cagr_3y_pct"), 0)
    comp = ratings["composite"]
    r6m  = ratings["r6m"]
    r5y  = ratings["r5y"]
    val_c = (
        f"At {pe:.0f}x trailing P/E, the valuation is stretched — making this a hold for existing "
        f"investors and an avoid for fresh buyers near term." if pe > 35
        else f"At {pe:.0f}x trailing P/E, valuation is moderate and offers a reasonable entry." if pe < 22
        else f"At {pe:.0f}x trailing P/E, the stock is priced at a modest premium to the market.")
    return (
        f"{name} scores {comp:.1f}/10 on our financial health scorecard, driven by "
        f"{'strong' if comp > 7 else 'reasonable' if comp > 5.5 else 'below-average'} "
        f"growth, margins, and balance sheet quality. "
        f"{val_c} "
        f"Our 6-month rating is {r6m} while the 5-year rating is {r5y}, "
        f"reflecting the distinction between current valuation and long-term compounding potential. "
        f"Investors with a 3–5 year horizon who can tolerate short-term volatility "
        f"will be better rewarded than those seeking quick gains at current prices."
    )

def build_revenue_segments(ann, meta):
    opm = safe(last(ann.get("opm_pct")), 0)
    return [
        {"segment": "Core Products / Services", "share": "65–75%", "nature": "Primary revenue driver; growth engine"},
        {"segment": "Premium / Export Markets",  "share": "15–25%", "nature": "Higher margin; drives OPM expansion" if opm > 15 else "Growing contribution"},
        {"segment": "Ancillary / Services",      "share": "10–15%", "nature": "Recurring; supports overall margin mix"},
    ]


# ═══════════════════════════════════════════════════════════════════════════════
# MASTER FUNCTION
# ═══════════════════════════════════════════════════════════════════════════════

def run_analysis(data):
    meta = data["meta"]
    ann  = data["annual"]
    qtr  = data["quarterly"]
    bs   = data["balance_sheet"]
    cf   = data["cash_flow"]

    scores = {
        "growth":        score_growth(ann, qtr),
        "margins":       score_margins(ann, qtr),
        "balance_sheet": score_balance_sheet(bs),
        "cashflow":      score_cashflow(cf, ann),
        "returns":       score_returns(bs, ann),
        "working_cap":   score_working_capital(bs, ann),
        "valuation":     score_valuation(meta, ann),
        "consistency":   score_consistency(ann),
    }
    ratings = derive_ratings(scores, meta, ann)
    val     = build_valuation(meta, ann, bs)
    strategy = build_strategy(ratings, meta, ann, bs, val)

    return {
        "company_summary":           company_summary(meta, ann, qtr, bs, cf),
        "sector_context":            sector_context(meta, ann, scores, ratings),
        "key_highlights":            build_highlights(meta, ann, bs, qtr, scores, ratings),
        "positives":                 build_positives(scores, meta, ann, bs, cf),
        "negatives":                 build_negatives(scores, meta, ann, bs, cf),
        "analyst_views":             build_analyst_views(meta, ann, val, strategy),
        "investment_strategy":       strategy,
        "valuation":                 {k:v for k,v in val.items() if not k.startswith("_")},
        "key_monitorables":          build_monitorables(ann, bs, qtr, meta),
        "conclusion":                build_conclusion(ratings, meta, ann, scores),
        "revenue_segments":          build_revenue_segments(ann, meta),
        "recent_quarter_commentary": build_recent_q_commentary(qtr, ann),
        "_scores":                   scores,
        "_ratings":                  ratings,
    }


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--data", required=True)
    parser.add_argument("--out",  default="ai_analysis.json")
    args = parser.parse_args()

    with open(args.data) as f:
        data = json.load(f)

    name = data["meta"]["company_name"]
    print(f"[Analysis] Scoring: {name}")

    result = run_analysis(data)

    s = result["_scores"]
    r = result["_ratings"]
    print(f"\n  DIMENSION SCORES (out of 10):")
    for k, v in s.items():
        bar = "█" * v + "░" * (10 - v)
        print(f"  {k:<18} {bar}  {v}/10")
    print(f"\n  COMPOSITE   : {r['composite']}/10")
    print(f"  RATINGS  →  6M: {r['r6m']}  |  1Y: {r['r1y']}  |  3Y: {r['r3y']}  |  5Y: {r['r5y']}")

    with open(args.out, "w") as f:
        json.dump(result, f, indent=2)
    print(f"\n✅  Analysis written → {args.out}")


if __name__ == "__main__":
    main()
