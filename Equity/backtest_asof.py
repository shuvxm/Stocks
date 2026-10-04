"""Point-in-time (no look-ahead) backtest of nifty_swing_analyst_v2.

Question answered: "If I had run the strategy on <asof> after the close, using only data up to
that day, would it have picked the stocks that later appeared in the alerts?"

Put this file next to nifty_swing_analyst_v2.py, then:

    python backtest_asof.py --asof 2026-09-29
    python backtest_asof.py --asof 2026-09-30            # fairer 1-day-ahead test for the 1-Oct alerts
    python backtest_asof.py --asof 2026-09-29 --universe NIFTY200   # what the default scan would have seen

How look-ahead is prevented
  * Every stock's data is cut at <asof> BEFORE any indicator is computed (EMAs, ATR, RSI, ADX,
    breakout levels, historical hit-rate). Nothing after <asof> touches the signal.
  * News, fundamentals, earnings dates and live prices come from "today", so they are NOT used.
    The score here is the pure technical score (the live script can add about -11 to +7 points
    from news/fundamentals/earnings). Stocks within 3 points of the cut are flagged BORDERLINE.
  * Prices after <asof> are used only AFTER the picks are fixed, to see what happened next.
Remaining small biases: today's NSE index list is used (index membership changes slowly).
"""
import argparse
from pathlib import Path

import numpy as np
import pandas as pd

import indian_stock_swing_analyst as n

# ---- The alerts you received. Edit symbols if Yahoo/NSE symbol differs (check on finance.yahoo.com).
# buy_low/buy_high: "Buy above X" -> buy_low=X, buy_high=None | range -> both | "Buy at CMP" -> both None
ALERTS = {
    "Tourism Finance Corporation of India Limited":    dict(alert="30-Sep", buy_low=149,  buy_high=None, targets=[155, 160, 170], sl=140),
    "Tega Industries Limited":       dict(alert="01-Oct", buy_low=2200, buy_high=None, targets=[2250, 2300, 2400, 2500, 2600], sl=2020),
    "Sona BLW Precision Forgings Limite":   dict(alert="01-Oct", buy_low=830,  buy_high=840,  targets=[850, 880, 900, 950, 1000], sl=800),
    "Rane (Madras) Limited":        dict(alert="01-Oct", buy_low=1416, buy_high=None, targets=[1450, 1470, 1550, 1650], sl=1310),  # <-- put the real NSE symbol
    "Gujarat Kidney And Super Speciality Limited":       dict(alert="30-Sep", buy_low=190,  buy_high=192,  targets=[200, 215, 230, 240], sl=185),
    "Rolex Rings Limited": dict(alert="30-Sep", buy_low=None, buy_high=None, targets=[200, 210, 225, 250], sl=186),
}


def cut(d, asof):
    d = d.copy()
    idx = pd.DatetimeIndex(d.index)
    if idx.tz is not None:
        idx = idx.tz_localize(None)
    d.index = idx.normalize()
    return d[d.index <= asof], d[d.index > asof]


def regime_asof(past):
    if past is None or len(past) < n.MIN_HISTORY:
        return {"regime": "UNKNOWN", "price": np.nan, "r21": np.nan, "r63": np.nan}
    r = n.indicators(past).iloc[-1]
    s = sum([r.Close > r.EMA20, r.EMA20 > r.EMA50, r.EMA50 > r.EMA200, r.RSI14 >= 50, r.ADX14 >= 20])
    return {"regime": "BULL" if s >= 4 else "BEAR" if s <= 1 else "NEUTRAL",
            "price": float(r.Close), "r21": float(r.RET21), "r63": float(r.RET63)}


def fail_list(past, m):
    """Every individual reason a stock would not make the list (not just the first)."""
    x = n.indicators(past)
    r = x.iloc[-1]
    out = []
    bad = [k for k in n.REQ if pd.isna(r.get(k))]
    if bad:
        return [f"insufficient data for: {', '.join(bad[:4])}"], r
    av, val = x.Volume.tail(20).mean(), (x.Close * x.Volume).tail(20).mean()
    if av < n.MIN_AVG_VOLUME: out.append(f"illiquid: avg vol {av:,.0f} < {n.MIN_AVG_VOLUME:,}")
    if val < n.MIN_AVG_VALUE: out.append(f"illiquid: avg value Rs{val/1e7:.1f}Cr < Rs{n.MIN_AVG_VALUE/1e7:.0f}Cr/day")
    s = n.classify(r)
    if s == "WATCH":
        out.append(f"no setup pattern (close {r.Close:.1f}, 20D high {r.High20P:.1f}, RVOL {r.RVOL:.2f}, RSI {r.RSI14:.0f})")
    k = s if s != "WATCH" else "TREND CONTINUATION"
    if not (r.Close > r.EMA50 and r.Close > r.EMA200 and r.EMA20 > r.EMA50):
        out.append("trend: needs Close>EMA50, Close>EMA200, EMA20>EMA50")
    rmin = 45 if k == "EMA20 PULLBACK" else n.RSI_MIN
    if not (rmin <= r.RSI14 <= n.RSI_MAX): out.append(f"RSI {r.RSI14:.0f} outside {rmin}-{n.RSI_MAX}")
    amin = 15 if k in ("BREAKOUT", "NEAR BREAKOUT") else n.ADX_MIN
    if r.ADX14 < amin: out.append(f"ADX {r.ADX14:.0f} < {amin}")
    if r.RVOL < n.RVOL_MIN.get(k, 1.0): out.append(f"RVOL {r.RVOL:.2f} < {n.RVOL_MIN.get(k, 1.0)}")
    if not (n.ATR_PCT_MIN <= r.ATRpct <= n.ATR_PCT_MAX): out.append(f"ATR% {r.ATRpct:.1f} outside {n.ATR_PCT_MIN}-{n.ATR_PCT_MAX}")
    if r.DistEMA20 > n.MAX_EXT_EMA20 or r.RET1 > 8 or r.RET5 > 15:
        out.append(f"extended/spiking (EMA20 dist {r.DistEMA20:.1f}%, 1D {r.RET1:.1f}%, 5D {r.RET5:.1f}%)")
    if m["regime"] == "BEAR" and np.isfinite(m["r21"]) and r.RET21 < m["r21"] + 5:
        out.append("bear market and not outperforming NIFTY")
    if not out:
        if n.trade_levels(r, k) is None: out.append("reward:risk too low")
    return out, r


def outcome(f, lo, hi, stop, t1, t2, kind):
    """What happened after the signal. kind='above' (buy only above lo) or 'zone' (buy next open if within zone)."""
    if f is None or f.empty:
        return {"Entry": "no later data"}
    if kind == "above":
        trig = f[f.High >= lo]
        if trig.empty:
            return {"Entry": "trigger never hit", "Max high after signal %": np.nan}
        i = trig.index[0]
        entry = max(lo, float(f.loc[i, "Open"]))
    else:
        o = float(f.Open.iloc[0])
        top = hi if (hi is not None and np.isfinite(hi)) else np.inf
        if o > top * 1.01:
            return {"Entry": "gapped above zone - skipped"}
        i, entry = f.index[0], o
    g = f.loc[i:]
    t1i = g.index[g.High >= t1][0] if (g.High >= t1).any() else None
    sli = g.index[g.Low <= stop][0] if (g.Low <= stop).any() else None
    if t1i is not None and (sli is None or t1i < sli): res = "T1 HIT"
    elif sli is not None and (t1i is None or sli < t1i): res = "STOP HIT"
    elif t1i is not None: res = "same-day T1/stop (ambiguous)"
    else: res = "open (neither yet)"
    return {"Entry": f"{entry:.2f}", "Result": res, "Bars held": len(g),
            "Max gain %": (g.High.max() / entry - 1) * 100, "Max drawdown %": (g.Low.min() / entry - 1) * 100,
            "Reached T2": bool((g.High >= t2).any())}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--asof", default="2026-09-29")
    ap.add_argument("--universe", choices=list(n.NSE_URLS), default="NIFTY500")
    ap.add_argument("--extra", default=",".join(ALERTS), help="comma list of NSE symbols to always evaluate")
    ap.add_argument("--period", default="3y")
    ap.add_argument("--min-score", type=int, default=n.MIN_SCORE)
    args = ap.parse_args()
    asof = pd.Timestamp(args.asof).normalize()
    n.HISTORY_PERIOD = args.period
    n.MIN_SCORE = args.min_score
    focus = [s.strip().upper() for s in args.extra.split(",") if s.strip()]

    print(f"POINT-IN-TIME BACKTEST | as of close {asof:%d-%b-%Y} | universe {args.universe}")
    tickers, sectors, names = n.get_universe(args.universe)
    base = set(tickers)
    all_t = list(tickers) + [s + ".NS" for s in focus if s + ".NS" not in base]
    data = n.download_all(all_t)
    print(f"\nData for {len(data)}/{len(all_t)} tickers")

    nf = n.history_single(n.NIFTY)
    npast, _ = cut(nf, asof) if nf is not None else (None, None)
    m = regime_asof(npast)
    print("NIFTY regime on that date:", m["regime"])

    rows, rejects, futs, pasts = {}, {}, {}, {}
    from collections import Counter
    funnel = Counter()
    for t, d in data.items():
        past, fut = cut(d, asof)
        futs[t] = fut
        if len(past) == 0:
            rejects[t] = "no data on/before asof"
            continue
        pasts[t] = past
        row, why = n.evaluate(t, past, m, sectors, names)
        if row:
            rows[t] = row
        else:
            rejects[t] = why
            if t in base:
                funnel[why] += 1

    cands = sorted((r for t, r in rows.items() if t in base), key=lambda r: r["Base Score"], reverse=True)
    picks = [r for r in cands if r["Base Score"] >= n.MIN_SCORE][:n.MAX_PICKS]
    pid = {id(r) for r in picks}
    watch = [r for r in cands if id(r) not in pid and r["Base Score"] >= n.WATCH_SCORE][:n.MAX_WATCH]
    pick_rank = {r["Stock"]: i for i, r in enumerate(picks, 1)}
    watch_rank = {r["Stock"]: i for i, r in enumerate(watch, 1)}
    scores = [r["Base Score"] for r in cands]

    # ---------------- focus stocks
    focus_rows = []
    for S in focus:
        t = S + ".NS"
        a = ALERTS.get(S, {})
        rec = {"Stock": S, "Alert date": a.get("alert", ""), "In scanned universe?": "Yes" if t in base else "NO - not in " + args.universe}
        fut = futs.get(t)
        if t not in data:
            rec.update({"Status on asof": "NO DATA (check Yahoo symbol)", "Why": ""})
        elif t in rows:
            r = rows[t]
            rank = 1 + sum(1 for s in scores if s > r["Base Score"])
            if S in pick_rank: st = f"PICK #{pick_rank[S]}"
            elif S in watch_rank: st = f"WATCHLIST #{watch_rank[S]}"
            elif r["Base Score"] >= n.MIN_SCORE: st = f"WOULD BE PICK (rank ~#{rank}, outside top {n.MAX_PICKS})"
            elif r["Base Score"] >= n.MIN_SCORE - 3: st = "BORDERLINE (within 3 pts of pick cut)"
            elif r["Base Score"] >= n.WATCH_SCORE: st = "WOULD BE WATCHLIST"
            else: st = f"passed filters, score {r['Base Score']} below watch cut"
            if t not in base: st += " [if it were scanned]"
            rec.update({"Status on asof": st, "Setup": r["Setup"], "Base Score": r["Base Score"], "Close asof": r["Last Close"],
                        "Strategy entry": f"{r['Entry Low']:.2f}-{r['Entry High']:.2f}", "Strategy SL": r["Stop Loss"],
                        "Strategy T1": r["Target 1"], "Why": " | ".join(r["_why"][:8])})
        else:
            why = rejects.get(t, "")
            detail = ""
            if t in pasts and len(pasts[t]) >= n.MIN_HISTORY:
                fl, r = fail_list(pasts[t], m)
                detail = " ; ".join(fl)
                rec["Close asof"] = float(r.Close) if "Close" in r else np.nan
            rec.update({"Status on asof": f"REJECTED - {why}", "Why": detail})
        if t in pasts and len(pasts[t]) >= n.MIN_HISTORY:
            x = n.indicators(pasts[t]); r = x.iloc[-1]
            rec.update({"RSI": r.RSI14, "ADX": r.ADX14, "RVOL": r.RVOL, "ATR%": r.ATRpct,
                        "EMA20 dist %": r.DistEMA20, "vs 20D high %": r.Dist20, "52W high dist %": r.Dist52W})
            rec["Avg value Rs Cr/day"] = float((x.Close * x.Volume).tail(20).mean() / 1e7)
        # what happened to the ALERT levels
        if a and fut is not None and len(fut):
            kind = "above" if (a["buy_low"] and a["buy_high"] is None) else "zone"
            lo = a["buy_low"] if a["buy_low"] else np.nan
            hi = a["buy_high"] if a["buy_high"] else (np.inf if kind == "zone" else None)
            oc = outcome(fut, lo, hi, a["sl"], a["targets"][0], a["targets"][1], kind)
            rec.update({f"Alert {k}": v for k, v in oc.items()})
            rec["Gain asof close -> best high %"] = (fut.High.max() / float(pasts[t].Close.iloc[-1]) - 1) * 100 if t in pasts else np.nan
        focus_rows.append(rec)

    # ---------------- base-rate: how did ALL picks do?
    pick_rows = []
    for r in picks + watch:
        fut = futs[r["_ticker"]]
        kind = "above" if r["Setup"] == "NEAR BREAKOUT" else "zone"
        oc = outcome(fut, r["Entry Low"], r["Entry High"], r["Stop Loss"], r["Target 1"], r["Target 2"], kind)
        pick_rows.append({"Tier": "PICK" if id(r) in pid else "WATCH", "Stock": r["Stock"], "Setup": r["Setup"],
                          "Base Score": r["Base Score"], "Close asof": r["Last Close"], "Entry Low": r["Entry Low"],
                          "Entry High": r["Entry High"], "Stop Loss": r["Stop Loss"], "Target 1": r["Target 1"], **oc})
    univ_gain = [(futs[t].High.max() / pasts[t].Close.iloc[-1] - 1) * 100 for t in pasts if t in base and len(futs[t])]
    nbars = int(max((len(futs[t]) for t in futs), default=0))

    # ---------------- print + save
    F = pd.DataFrame(focus_rows)
    pd.set_option("display.width", 220, "display.max_columns", 30, "display.max_colwidth", 70)
    print("\n================ ALERT STOCKS: WHAT THE STRATEGY SAW ON", f"{asof:%d-%b-%Y} ================")
    show = [c for c in ["Stock", "In scanned universe?", "Status on asof", "Setup", "Base Score"] if c in F]
    print(F[show].to_string(index=False))
    st = F["Status on asof"].astype(str)
    hit = st.str.startswith(("PICK", "WATCHLIST", "WOULD BE PICK", "BORDERLINE", "WOULD BE WATCHLIST")).sum()
    print(f"\nCaught (pick/watch/borderline or would-be): {hit} of {len(F)}")
    print(f"Strictly in SWING PICKS: {st.str.startswith('PICK').sum()} | not even in scanned universe: {(~F['In scanned universe?'].eq('Yes')).sum()}")
    print("\nRejection reasons for the missed ones:")
    for _, r in F.iterrows():
        if str(r["Status on asof"]).startswith(("REJECTED", "NO DATA")):
            print(f"  {r['Stock']}: {r['Status on asof']}\n      {r.get('Why', '')}")
    print(f"\nBase rate: {len(picks)} picks / {len(watch)} watch on that date. Bars of later data available: {nbars}")
    if pick_rows:
        P = pd.DataFrame(pick_rows)
        pk = P[P.Tier == "PICK"]
        if len(pk) and "Result" in pk:
            print("Picks outcome:", pk["Result"].value_counts(dropna=False).to_dict(),
                  "| median max gain %:", round(pk["Max gain %"].median(), 1))
    if univ_gain:
        print(f"Whole universe median best-high gain over same window: {np.nanmedian(univ_gain):.1f}% "
              f"| share of stocks that touched +5%: {np.mean(np.array(univ_gain) >= 5)*100:.0f}%")
    print("NOTE: with only a few bars after the signal date, this is a tiny sample. Run several --asof dates to judge the strategy.")

    out = Path(__file__).resolve().parent / "output" / "daily"
    out.mkdir(parents=True, exist_ok=True)
    path = out / f"Backtest_asof_{asof:%Y-%m-%d}.xlsx"
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        F.to_excel(w, sheet_name="ALERT STOCKS", index=False)
        pd.DataFrame(pick_rows).to_excel(w, sheet_name="PICKS ON DATE", index=False)
        pd.DataFrame(sorted(funnel.items(), key=lambda kv: -kv[1]), columns=["Rejection reason", "Stocks"]).to_excel(w, sheet_name="FUNNEL", index=False)
        for ws in w.book.worksheets:
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = min(60, max(12, max(len(str(c.value or "")) for c in col[:30]) + 2))
    print("\nSaved:", path)


if __name__ == "__main__":
    main()