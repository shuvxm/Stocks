"""Test the swing scanner on the last 6 months, one day at a time (no look-ahead).

For every trading day in the window this re-runs the scanner's own rules using ONLY the prices
known up to that day's close, takes the picks it would have reported that evening, and then
follows each pick with the same plan as backtest_picks.py (entry zone on the next session,
then Target 1 or Stop Loss, whichever comes first).

Install : pip install yfinance pandas numpy openpyxl requests
Run     : python backtest_6months.py                      (same settings as the evening scan:
                                                           NIFTY500, smallcap mode, 15 picks, score 60+)
Options : --months 6            how far back to test
          --universe NIFTY500   NIFTY50 | NIFTY100 | NIFTY200 | NIFTY500 | ...
          --strict              strict filters only (no smallcap 2nd pass)
          --picks 15            picks per day        --min-score 60   minimum score for a pick
Output  : output/Backtest6Months_<Sun_11th Oct>.xlsx   (same columns as the picks backtest)
Sheets  : PICKS RESULT | SUMMARY (by month) | BY SETUP | BY SCORE | BY NIFTY REGIME | NOTES

What is and is not the same as the live scanner
  * Same code: indicators, setups, filters, score, entry / stop / targets, quantity.
  * NOT included: news, fundamentals, earnings date and live price. Old news and old fundamentals
    cannot be fetched, so the score here is the technical "Base Score" (the live report moves it
    by about -11 to +7 points and can drop a stock on bad news).
  * Today's index list is used for every past day (stocks that left the index are missing).
  * A stock is not bought again while an earlier pick of it is still open.
"""
import argparse
import sys
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))   # the scanner lives in Equity/
import indian_stock_swing_analyst as n

import backtest_picks as bp

OHLC = ["Open", "High", "Low", "Close"]


def load(universe, period):
    """Download daily prices for the universe + NIFTY, with plain (no time zone) dates."""
    n.HISTORY_PERIOD = period
    tickers, sectors, names = n.get_universe(universe)
    print("Stocks in universe:", len(tickers))
    data = n.download_all(tickers)
    print(f"\nPrice data received for {len(data)}/{len(tickers)} stocks")
    nifty = n.history_single(n.NIFTY)
    if nifty is None:
        sys.exit("Could not download NIFTY prices.")

    def tidy(d):
        d = n.drop_partial(d).copy()
        idx = pd.DatetimeIndex(d.index)
        d.index = (idx.tz_localize(None) if idx.tz is not None else idx).normalize()
        return d[~d.index.duplicated(keep="last")]

    return {t: tidy(d) for t, d in data.items()}, tidy(nifty), sectors, names


class IndicatorCache:
    """The scanner recomputes every indicator from scratch on each call. All of them only look
    backwards, so the values for "data up to day k" are simply the first k rows of the values for
    the full history. Computing once per stock makes a 6-month test take minutes instead of an hour.
    The one forward-looking column (FwdMaxPct, used for the historical hit-rate) is blanked for the
    last HOLD_DAYS rows of each cut, exactly as it would be if the later prices did not exist."""

    def __init__(self, data):
        self.real = n.indicators
        self.full = {t: self.real(d) for t, d in data.items()}
        self.fwd = {t: x.columns.get_loc("FwdMaxPct") for t, x in self.full.items()}
        self.current = None

    def __call__(self, d):
        x = self.full[self.current].iloc[:len(d)].copy()
        x.iloc[max(0, len(x) - n.HOLD_DAYS):, self.fwd[self.current]] = np.nan
        return x

    def __enter__(self):
        n.indicators = self
        return self

    def __exit__(self, *a):
        n.indicators = self.real


def regime_on(nx, asof):
    """NIFTY regime exactly as the scanner computes it, from rows up to asof (nx = NIFTY indicators)."""
    p = nx.loc[:asof]
    if len(p) < n.MIN_HISTORY:
        return {"regime": "UNKNOWN", "price": np.nan, "r21": np.nan, "r63": np.nan}
    r = p.iloc[-1]
    s = sum([r.Close > r.EMA20, r.EMA20 > r.EMA50, r.EMA50 > r.EMA200, r.RSI14 >= 50, r.ADX14 >= 20])
    return {"regime": "BULL" if s >= 4 else "BEAR" if s <= 1 else "NEUTRAL",
            "price": float(r.Close), "r21": float(r.RET21), "r63": float(r.RET63)}


def scan_day(asof, data, cache, m, sectors, names, fn, min_score, max_picks):
    """The picks the scanner would have reported after the close of asof."""
    cands = []
    for t, d in data.items():
        k = int(d.index.searchsorted(asof, side="right"))
        if k == 0 or d.index[k - 1] != asof:       # no candle that day (not listed yet / not traded)
            continue
        cache.current = t
        try:
            row, _ = fn(t, d.iloc[:k], m, sectors, names)
        except Exception:
            row = None
        if row:
            row["_k"] = k
            cands.append(row)
    cands.sort(key=lambda r: r["Base Score"], reverse=True)
    return [r for r in cands if r["Base Score"] >= min_score][:max_picks]


def peak_money(res):
    """Most money tied up in open trades on any one day, if every pick was taken."""
    ent = res[res["Entry Hit?"] == "YES"]
    if ent.empty:
        return 0.0
    ev = [(pd.Timestamp(r["Purchase Date"]), r["Invested Rs"]) for _, r in ent.iterrows()]
    ev += [(pd.Timestamp(r["Exit Date"]) + pd.Timedelta(days=1), -r["Invested Rs"])
           for _, r in ent.iterrows() if pd.notna(r["Exit Date"])]
    run = peak = 0.0
    for _, amt in sorted(ev, key=lambda e: (e[0], e[1])):
        run += amt
        peak = max(peak, run)
    return peak


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--months", type=int, default=6)
    ap.add_argument("--universe", choices=list(n.NSE_URLS), default="NIFTY500")
    ap.add_argument("--strict", action="store_true", help="strict filters only (no smallcap 2nd pass)")
    ap.add_argument("--picks", type=int, default=n.MAX_PICKS)
    ap.add_argument("--min-score", type=int, default=n.MIN_SCORE)
    ap.add_argument("--period", default="3y", help="price history to download (must cover the test + 1 year)")
    args = ap.parse_args()
    fn = n.evaluate if args.strict else n.evaluate_relaxed

    print(f"SCANNER BACKTEST | last {args.months} months | universe {args.universe} | "
          f"{'strict' if args.strict else 'smallcap mode'} | top {args.picks} picks/day, score {args.min_score}+")
    data, nifty, sectors, names = load(args.universe, args.period)
    last = nifty.index[-1]
    start = last - pd.DateOffset(months=args.months)
    days = nifty.index[(nifty.index >= start) & (nifty.index < last)]   # the last day has no next session yet
    print(f"Scan days: {len(days)}  ({days[0]:%d-%b-%Y} to {days[-1]:%d-%b-%Y}) | prices up to {last:%d-%b-%Y}")

    rows, busy, repeats, regimes = [], {}, 0, []
    nx = n.indicators(nifty)
    with IndicatorCache(data) as cache:
        for i, asof in enumerate(days, 1):
            m = regime_on(nx, asof)
            regimes.append(m["regime"])
            for rank, r in enumerate(scan_day(asof, data, cache, m, sectors, names, fn,
                                              args.min_score, args.picks), 1):
                d = data[r["_ticker"]]
                fut = d.iloc[r["_k"]:]
                if fut.empty:
                    continue
                buy_day = fut.index[0].date()
                if r["Stock"] in busy and buy_day <= busy[r["Stock"]]:
                    repeats += 1                    # still holding it from an earlier pick
                    continue
                pick = {"Stock": r["Stock"], "Setup": r["Setup"], "Swing Score": r["Base Score"],
                        "Entry Low": r["Entry Low"], "Entry High": r["Entry High"], "Stop Loss": r["Stop Loss"],
                        "Target 1": r["Target 1"], "Target 2": r["Target 2"], "Qty (report)": r["Suggested Qty*"],
                        "Report": f"Scan of {asof:%d-%b-%Y} (rank {rank})",
                        "Generated": asof + pd.Timedelta(hours=16, minutes=30)}
                res = bp.evaluate(pick, fut[OHLC], float(d["Close"].iloc[-1]), f"Close {d.index[-1]:%d-%b}")
                res["NIFTY Regime"] = m["regime"]
                if res["Entry Hit?"] == "YES":
                    busy[r["Stock"]] = res["Exit Date"] or datetime.max.date()
                rows.append(res)
            print(f"  scanned {i}/{len(days)} days, {len(rows)} picks so far", end="\r")
    if not rows:
        sys.exit("\nThe scanner gave no picks in this period.")

    res = pd.DataFrame(rows)
    res["Month"] = [f"{d:%b %Y}" for d in res["Purchase Date"]]
    res["Score Band"] = pd.cut(res["Swing Score"], [0, 64, 69, 74, 100], labels=["60-64", "65-69", "70-74", "75+"]).astype(str)
    by_month = bp.summary(res, "Month", "Month")
    by_setup = bp.summary(res, "Setup", "Setup")
    by_score = bp.summary(res.sort_values("Swing Score", kind="stable"), "Score Band", "Score")
    by_regime = bp.summary(res, "NIFTY Regime", "NIFTY regime on scan day")

    n0, n1 = float(nifty["Close"].loc[days[0]]), float(nifty["Close"].iloc[-1])
    peak = peak_money(res)
    reg = pd.Series(regimes).value_counts().to_dict()
    notes = [
        f"SCANNER BACKTEST: {len(days)} trading days, scans of {days[0]:%d-%b-%Y} to {days[-1]:%d-%b-%Y}, "
        f"prices up to {last:%d-%b-%Y}. Universe {args.universe}, {'strict filters' if args.strict else 'smallcap mode'}, "
        f"top {args.picks} picks a day with score {args.min_score}+ (same as the evening scan).",
        "Each day the scanner's own rules were re-run using only prices up to that day's close, so no pick could "
        "see the future. Purchase Date is the next market session.",
        "NOT the same as the live report: news, fundamentals, earnings dates and live price are not used, because old "
        "news and old fundamentals cannot be fetched. The score is the technical Base Score; the live report moves it "
        "by about -11 to +7 points and can drop a stock on bad news.",
        "Today's index list is used for every past day, so stocks that have left the index are missing (this usually "
        "makes a backtest look a little better than real life).",
        f"A stock is not bought again while an earlier pick of it is still open: {repeats} repeat picks were left out.",
        f"NIFTY over the same period: {n0:,.0f} to {n1:,.0f} ({(n1 / n0 - 1) * 100:+.1f}%). "
        f"NIFTY regime on the scan days: {reg}.",
        f"Money: quantity is sized for Rs 1L capital per trade idea, but taking EVERY pick needed up to "
        f"Rs {peak:,.0f} tied up at one time. Total Gain/Loss Rs is the sum over all trades, not the return on Rs 1L.",
    ] + bp.NOTES[1:]

    bp.OUT_DIR.mkdir(parents=True, exist_ok=True)
    path = bp.write(res, bp.OUT_DIR / f"Backtest6Months_{bp.day_tag(datetime.now())}.xlsx",
                    {"SUMMARY": by_month, "BY SETUP": by_setup, "BY SCORE": by_score, "BY NIFTY REGIME": by_regime},
                    notes)

    pd.set_option("display.width", 250, "display.max_columns", 30)
    fmt = lambda v: f"{v:,.1f}"
    print("\n\nBY MONTH (month of purchase)")
    print(by_month.drop(columns=["Best", "Worst"], errors="ignore").to_string(index=False, float_format=fmt))
    print("\nBY SETUP")
    print(by_setup.drop(columns=["Best", "Worst"], errors="ignore").to_string(index=False, float_format=fmt))
    print("\nBY SCORE")
    print(by_score.drop(columns=["Best", "Worst"], errors="ignore").to_string(index=False, float_format=fmt))
    print("\nBY NIFTY REGIME")
    print(by_regime.drop(columns=["Best", "Worst"], errors="ignore").to_string(index=False, float_format=fmt))
    print(f"\nRepeat picks left out (already holding): {repeats}")
    print(f"NIFTY same period: {(n1 / n0 - 1) * 100:+.1f}% | regime days: {reg}")
    print(f"Most money tied up at one time if every pick was taken: Rs {peak:,.0f}")
    print("\nSaved:", path)


if __name__ == "__main__":
    main()
