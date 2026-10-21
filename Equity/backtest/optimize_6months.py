"""Test scanner improvements on the last 6 months, one change at a time (no look-ahead).

Step 1 (slow, once): every trading day the scanner's own rules are re-run on prices known up to that
close (same as backtest_6months.py) and EVERY candidate is recorded with the extra facts the changes need.
Step 2 (fast): each variant = the current rules + one change. It picks its SWING PICKS (top 15) and
PRIORITY stocks from the recorded candidates, follows each trade on real daily candles, and runs a
Rs 1,00,000 portfolio on the PRIORITY picks (a stock is bought only if there is cash for it).

Install : pip install yfinance pandas numpy openpyxl requests
Run     : python optimize_6months.py --cache <folder>     (prices and candidates are saved there, so a
                                                           second run only re-does step 2)
Output  : ../../REPORTS/Optimization_6Months_<Wed_21st Oct>.xlsx

Variants (see VARIANTS below)
  BASE        the scanner as it is today
  HIST_FIX    Hist Hit 7% counts a hit only if +7% came BEFORE the stop (today it counts any +7% in 15 days)
  FLAGS_FIX   smallcap ("RELAXED") picks get the same risk flags as the others (before 21-Oct-2026 they got
              none, so an overheated smallcap passed the PRIORITY "no Risk Flags" check; the scanner now
              does this, so from then on FLAGS_FIX = BASE)
  NAT_STOP    skip a stock whose normal stop (1.8 ATR / 10-day low) is more than 6.5% away, instead of
              squeezing its stop to 6.5%
  EXT7        skip a stock more than 7% above its EMA20 (today 10%, smallcap 12%)
  STRICT      no smallcap 2nd pass
  NO_BREAKOUT skip the BREAKOUT setup (stock already ran past its 20-day high)
  BEAR_LIGHT  on a BEAR day: 5 swing picks and 1 PRIORITY stock only
  BEAR_OFF    on a BEAR day: no new trades
  HALF_EXIT   sell half at Target 1, move the stop of the other half to the buy price, sell it at Target 2
  T5          Target 1 = +5%, Target 2 = +7% (stop unchanged)
"""
import argparse
import pickle
import sys
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))   # the scanner lives in Equity/
import indian_stock_swing_analyst as n

import backtest_6months as b6
import backtest_picks as bp

BASE = Path(__file__).resolve().parent
OUT_DIR = BASE.parents[1] / "REPORTS"
OHLC = ["Open", "High", "Low", "Close"]
CAPITAL = n.ACCOUNT_CAPITAL
MIN_RECORD = 50                     # candidates below this score cannot reach 60 under any variant

VARIANTS = {
    "BASE": {},
    "HIST_FIX": {"hist_fix": True},
    "FLAGS_FIX": {"flags_fix": True},
    "NAT_STOP": {"nat_stop": True},
    "EXT7": {"max_ext": 7.0},
    "STRICT": {"strict": True},
    "NO_BREAKOUT": {"no_breakout": True},
    "BEAR_LIGHT": {"bear": "light"},
    "BEAR_OFF": {"bear": "off"},
    "HALF_EXIT": {"exit": "half"},
    "T5": {"t1": 5.0, "t2": 7.0},
}
DESCRIBE = {
    "BASE": "Scanner exactly as it is today",
    "HIST_FIX": "Hist Hit 7% counts +7% only if it came before the stop",
    "FLAGS_FIX": "Smallcap picks get real risk flags (hot RSI, extended, spike...)",
    "NAT_STOP": "Skip stocks whose normal stop is more than 6.5% away",
    "EXT7": "Skip stocks more than 7% above EMA20 (don't chase)",
    "STRICT": "No smallcap (relaxed) picks",
    "NO_BREAKOUT": "Skip BREAKOUT setup (train already left)",
    "BEAR_LIGHT": "BEAR day: only 5 swing picks, 1 PRIORITY",
    "BEAR_OFF": "BEAR day: no new trades",
    "HALF_EXIT": "Sell half at T1, stop to buy price, rest at T2",
    "T5": "Target 1 = +5%, Target 2 = +7%",
}


# ----------------------------------------------------------------------------
# Step 1: record every candidate
# ----------------------------------------------------------------------------
def fwd_win(x):
    """Per day: 1 if +TARGET1_PCT was reached before the stop within HOLD_DAYS (stop as the scanner sets it:
    1.8 ATR / 10-day low, kept 3-6.5% below), 0 if the stop came first or neither, NaN without enough future."""
    c, h, l = (x[k].to_numpy(float) for k in ("Close", "High", "Low"))
    stop = np.maximum(c - n.STOP_ATR * x["ATR14"].to_numpy(float), x["Low10"].to_numpy(float) * 0.995)
    stop = np.maximum(np.minimum(stop, c * (1 - n.STOP_MIN_PCT / 100)), c * (1 - n.STOP_MAX_PCT / 100))
    tgt, N, H = c * (1 + n.TARGET1_PCT / 100), len(c), n.HOLD_DAYS
    first_t, first_s = np.full(N, np.inf), np.full(N, np.inf)
    for k in range(min(H, N - 1), 0, -1):         # walk back so the earliest day wins
        hk, lk = np.full(N, np.nan), np.full(N, np.nan)
        hk[:N - k], lk[:N - k] = h[k:], l[k:]
        first_t = np.where(hk >= tgt, k, first_t)
        first_s = np.where(lk <= stop, k, first_s)
    win = (first_t < first_s).astype(float)      # same day = stop first (conservative)
    win[max(0, N - H):] = np.nan
    win[~np.isfinite(stop)] = np.nan
    return win


class Cache(b6.IndicatorCache):
    """backtest_6months' indicator cache plus the FwdWin column, blanked the same way as FwdMaxPct."""

    def __init__(self, data):
        super().__init__(data)
        for x in self.full.values():
            x["FwdWin"] = fwd_win(x)
        self.win = {t: x.columns.get_loc("FwdWin") for t, x in self.full.items()}

    def __call__(self, d):
        x = super().__call__(d)
        x.iloc[max(0, len(x) - n.HOLD_DAYS):, self.win[self.current]] = np.nan
        return x


STASH = {}


def hist_both(real):
    """Wrap the scanner's hist_stats: same result, and also stash the stop-aware hit-rate."""
    def f(x):
        out = real(x)
        mask = ((x.Close > x.EMA50) & (x.EMA20 > x.EMA50) & (x.EMA50 > x.EMA200)
                & x.RSI14.between(50, 78) & (x.RVOL >= 0.9))
        w = x.FwdWin[mask].dropna()
        STASH["h7"] = w.mean() * 100 if len(w) >= 10 else np.nan
        return out
    return f


def hh(h7):
    """Hist-hit points of the scanner's score()."""
    return 3 if not np.isfinite(h7) else 10 if h7 >= 35 else 7 if h7 >= 25 else 4 if h7 >= 15 else 0


def record(data, nifty, sectors, names, days):
    nx = n.indicators(nifty)
    rows, real_hist = [], n.hist_stats
    n.hist_stats = hist_both(real_hist)
    try:
        with Cache(data) as cache:
            for i, asof in enumerate(days, 1):
                m = b6.regime_on(nx, asof)
                for t, d in data.items():
                    k = int(d.index.searchsorted(asof, side="right"))
                    if k == 0 or d.index[k - 1] != asof:
                        continue
                    cache.current = t
                    STASH.clear()
                    try:
                        row, _ = n.evaluate_relaxed(t, d.iloc[:k], m, sectors, names)
                    except Exception:
                        row = None
                    if not row or row["Base Score"] < MIN_RECORD:
                        continue
                    r = cache.full[t].iloc[k - 1]
                    ref = (row["Entry Low"] + row["Entry High"]) / 2 if row["Setup"] == "NEAR BREAKOUT" else float(r.Close)
                    nat = max(ref - n.STOP_ATR * float(r.ATR14), float(r.Low10) * 0.995)
                    h7_old, h7_new = row["Hist Hit 7%"], STASH.get("h7", np.nan)
                    rows.append({
                        "Day": asof, "Ticker": t, "k": k, "Stock": row["Stock"], "Sector": row["Sector"],
                        "Setup": row["Setup"], "Mode": row.get("_mode", "strict"), "Regime": m["regime"],
                        "Score": row["Base Score"],
                        "Score New": int(max(0, min(100, row["Base Score"] - hh(h7_old) + hh(h7_new)))),
                        "H7 Old": h7_old, "H7 New": h7_new, "Ref": ref,
                        "Entry Low": row["Entry Low"], "Entry High": row["Entry High"], "Stop Loss": row["Stop Loss"],
                        "Target 1": row["Target 1"], "Target 2": row["Target 2"], "Qty": row["Suggested Qty*"],
                        "Nat Stop %": (ref - nat) / ref * 100, "EMA20 Dist%": float(r.DistEMA20),
                        "Action": n.entry_status(float(r.Close), row),
                        "Flags Old": [f for f in row["_flags"] if not f.startswith("RELAXED")],
                        "Flags Full": n.risk_flags(r, h7_old),
                    })
                print(f"  recorded {i}/{len(days)} days, {len(rows)} candidates", end="\r")
    finally:
        n.hist_stats = real_hist
    print()
    return pd.DataFrame(rows)


# ----------------------------------------------------------------------------
# Step 2: variants
# ----------------------------------------------------------------------------
def select(cands, cfg):
    """{day: (swing picks, priority picks)} for one variant."""
    c = cands
    if cfg.get("strict"):
        c = c[c["Mode"] == "strict"]
    if cfg.get("nat_stop"):
        c = c[c["Nat Stop %"] <= n.STOP_MAX_PCT]
    if cfg.get("max_ext"):
        c = c[c["EMA20 Dist%"] <= cfg["max_ext"]]
    if cfg.get("no_breakout"):
        c = c[c["Setup"] != "BREAKOUT"]
    sc, h7 = ("Score New", "H7 New") if cfg.get("hist_fix") else ("Score", "H7 Old")
    flags = "Flags Full" if cfg.get("flags_fix") else "Flags Old"
    c = c[c[sc] >= n.MIN_SCORE]
    out = {}
    for day, g in c.groupby("Day", sort=True):
        bear = g["Regime"].iloc[0] == "BEAR"
        if bear and cfg.get("bear") == "off":
            continue
        swing = g.sort_values(sc, ascending=False, kind="stable").head(5 if bear and cfg.get("bear") == "light" else n.MAX_PICKS)
        swing = swing.assign(Rank=range(1, len(swing) + 1))
        ok = swing[swing["Action"].str.startswith("BUY ZONE") & (swing[flags].str.len() == 0)
                   & swing[h7].notna() & (swing[h7] >= n.PRIORITY_MIN_HIT)]
        ok = ok.sort_values([h7, "Rank"], ascending=[False, True], kind="stable")
        limit = (1 if cfg.get("bear") == "light" else n.PRIORITY_MAX_BEAR) if bear else n.PRIORITY_MAX
        prio, taken = [], set()
        for _, r in ok.iterrows():
            sector = r["Sector"] if isinstance(r["Sector"], str) and r["Sector"].strip() else r["Stock"]
            if sector not in taken and len(prio) < limit:
                taken.add(sector)
                prio.append(r)
        out[day] = (swing, pd.DataFrame(prio))
    return out


def sim_half(bars, lo, hi, stop, t1, t2):
    """Half at Target 1, the other half's stop moves to the buy price, sold at Target 2.
    Returns (simulate() result, final exit bar or None, average exit price or None)."""
    s = bp.simulate(bars, lo, hi, stop, t1, t2)
    if s["result"] != "TARGET HIT":
        return s, s["exit"], s["exit_px"]
    o, h, l = (bars[c].to_numpy(float) for c in ("Open", "High", "Low"))
    i, buy = s["exit"], s["buy"]
    if h[i] >= t2:                                # ran through both targets that day
        return s, i, (s["exit_px"] + max(t2, o[i] if i > s["entry"] else t2)) / 2
    for j in range(i + 1, len(o)):
        if l[j] <= buy:                           # back to the buy price: second half out at no loss
            return s, j, (s["exit_px"] + min(buy, o[j])) / 2
        if h[j] >= t2:
            return s, j, (s["exit_px"] + max(t2, o[j])) / 2
    return s, None, (s["exit_px"] + float(bars["Close"].iloc[-1])) / 2


SIMS = {}


def trade(data, r, cfg):
    """Follow one pick. Cached: the same pick shows up in many variants."""
    t1p, t2p, plan = cfg.get("t1", n.TARGET1_PCT), cfg.get("t2", n.TARGET2_PCT), cfg.get("exit", "t1")
    key = (r["Ticker"], r["k"], t1p, t2p, plan)
    if key in SIMS:
        return SIMS[key]
    d = data[r["Ticker"]]
    bars = d.iloc[r["k"]:][OHLC]
    t1, t2 = r["Ref"] * (1 + t1p / 100), r["Ref"] * (1 + t2p / 100)
    out = {"Entered": False}
    if len(bars):
        if plan == "half":
            s, ex, px = sim_half(bars, r["Entry Low"], r["Entry High"], r["Stop Loss"], t1, t2)
        else:
            s = bp.simulate(bars, r["Entry Low"], r["Entry High"], r["Stop Loss"], t1, t2)
            ex, px = s["exit"], s["exit_px"]
        if s["entry"] is not None:
            buy, last = s["buy"], float(bars["Close"].iloc[-1])
            px = px if px is not None and np.isfinite(px) else last
            qty = int(r["Qty"]) if pd.notna(r["Qty"]) else 0
            out = {"Entered": True, "Buy Date": bars.index[s["entry"]], "Exit Date": bars.index[ex] if ex is not None else None,
                   "Buy": buy, "Qty": qty, "T1 Hit": s["result"] == "TARGET HIT", "T2 Hit": bool(s["t2"]),
                   "SL Hit": s["result"] == "STOP LOSS HIT", "Open": s["result"] == "OPEN",
                   "PnL %": (px / buy - 1) * 100, "PnL Rs": (px - buy) * max(qty, 1)}
    SIMS[key] = out
    return out


def follow(data, picks, cfg):
    """Trades for a list of picks; a stock is not bought again while the earlier trade is open."""
    busy, rows = {}, []
    for day in sorted(picks):
        for _, r in picks[day].iterrows():
            res = trade(data, r, cfg)
            if not res["Entered"]:
                rows.append({"Day": day, "Stock": r["Stock"], "Entered": False})
                continue
            if r["Stock"] in busy and res["Buy Date"] <= busy[r["Stock"]]:
                continue
            busy[r["Stock"]] = res["Exit Date"] or pd.Timestamp.max
            rows.append({"Day": day, "Stock": r["Stock"], "Setup": r["Setup"], "Regime": r["Regime"], **res})
    return pd.DataFrame(rows)


def portfolio(trades, days):
    """Rs 1L account on the PRIORITY picks: buy only with free cash (Suggested Qty), cash comes back on exit day."""
    if trades.empty:
        return {"Trades Taken": 0, "Skipped (no cash)": 0, "Account P&L Rs": 0.0, "P&L per day Rs": 0.0, "Max Drawdown Rs": 0.0}
    t = trades[trades["Entered"] & (trades["Qty"] > 0)].sort_values("Buy Date", kind="stable")
    cash, held, taken, skipped, curve = float(CAPITAL), [], 0, 0, []
    realized = 0.0
    for day in sorted(set(t["Buy Date"]) | set(t["Exit Date"].dropna())):
        for h in [h for h in held if h["Exit Date"] is not None and h["Exit Date"] < day]:
            cash += h["Buy"] * h["Qty"] + h["PnL Rs"]
            realized += h["PnL Rs"]
            held.remove(h)
        for _, r in t[t["Buy Date"] == day].iterrows():
            cost = r["Buy"] * r["Qty"]
            if cost > cash or any(h["Stock"] == r["Stock"] for h in held):
                skipped += 1
                continue
            cash -= cost
            held.append(r.to_dict())
            taken += 1
        curve.append(realized)
    final = realized + sum(h["PnL Rs"] for h in held)          # still open / closing on the last day
    peak, dd = 0.0, 0.0
    for v in curve + [final]:
        peak, dd = max(peak, v), min(dd, v - peak)
    return {"Trades Taken": taken, "Skipped (no cash)": skipped, "Account P&L Rs": final,
            "P&L per day Rs": final / len(days), "Max Drawdown Rs": dd}


def stats(t):
    e = t[t["Entered"]] if len(t) else t
    nE = len(e)
    if not nE:
        return {"Picks": len(t), "Entered": 0}
    gains, losses = e.loc[e["PnL %"] > 0, "PnL %"].sum(), -e.loc[e["PnL %"] < 0, "PnL %"].sum()
    closed = int(e["T1 Hit"].sum() + e["SL Hit"].sum())
    return {"Picks": len(t), "Entered": nE, "T1 Hit %": e["T1 Hit"].mean(), "T2 Hit %": e["T2 Hit"].mean(),
            "SL Hit %": e["SL Hit"].mean(), "Win % (closed)": e["T1 Hit"].sum() / closed if closed else np.nan,
            "Avg P&L %": e["PnL %"].mean() / 100, "Profit Factor": gains / losses if losses else np.nan,
            "Total P&L Rs (sum of trades)": e["PnL Rs"].sum()}


def run_variants(data, cands, days, variants):
    split = days[len(days) // 2]
    rows, trades = [], {}
    for name, cfg in variants.items():
        sel = select(cands, cfg)
        sw = follow(data, {d: v[0] for d, v in sel.items()}, cfg)
        pr = follow(data, {d: v[1] for d, v in sel.items() if len(v[1])}, cfg)
        trades[name] = (sw, pr)
        row = {"Variant": name, "What changes": DESCRIBE.get(name, name)}
        for label, t in (("SWING", sw), ("PRIORITY", pr)):
            for k, v in stats(t).items():
                row[f"{label} {k}"] = v
            for half, part in (("1st half", t[t["Day"] < split] if len(t) else t),
                               ("2nd half", t[t["Day"] >= split] if len(t) else t)):
                s = stats(part)
                row[f"{label} Avg P&L % {half}"] = s.get("Avg P&L %", np.nan)
        for k, v in portfolio(pr, days).items():
            row[f"Rs 1L {k}"] = v
        rows.append(row)
        print(f"  {name:<12} swing avg {row.get('SWING Avg P&L %', np.nan) * 100:+.2f}%  "
              f"priority avg {row.get('PRIORITY Avg P&L %', np.nan) * 100:+.2f}%  "
              f"Rs 1L account {row['Rs 1L Account P&L Rs']:+,.0f} ({row['Rs 1L P&L per day Rs']:+,.0f}/day)")
    return pd.DataFrame(rows), trades, split


# ----------------------------------------------------------------------------
# Excel
# ----------------------------------------------------------------------------
def write(table, trades, days, split, path, notes):
    from openpyxl import Workbook
    from openpyxl.formatting.rule import FormulaRule
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    F = "Arial"
    head_f, body_f = Font(name=F, size=9, bold=True, color="FFFFFF"), Font(name=F, size=9)
    title_f, sub_f = Font(name=F, size=14, bold=True, color="FFFFFF"), Font(name=F, size=9, italic=True, color="404040")
    title_fill, head_fill = PatternFill("solid", fgColor="1F4E78"), PatternFill("solid", fgColor="2E75B6")
    base_fill = PatternFill("solid", fgColor="DDEBF7")
    green, red = PatternFill("solid", fgColor="C6EFCE"), PatternFill("solid", fgColor="FFC7CE")
    thin = Side(style="thin", color="B4C6E7")
    bd = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)

    def fmt(c):
        if "%" in c and "Rs" not in c:
            return "0.0%" if not c.endswith("Stop %") else "0.0"
        if "Rs" in c:
            return "#,##0;[Red]-#,##0"
        if "Factor" in c:
            return "0.00"
        return "0"

    def sheet(ws, df, title, sub, r0=4):
        cols = list(df.columns)
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(2, len(cols)))
        ws.cell(1, 1, title).font, ws.cell(1, 1).fill = title_f, title_fill
        ws.row_dimensions[1].height = 26
        ws.cell(2, 1, sub).font = sub_f
        for j, c in enumerate(cols, 1):
            x = ws.cell(r0, j, c)
            x.font, x.fill, x.alignment, x.border = head_f, head_fill, center, bd
            ws.column_dimensions[get_column_letter(j)].width = 11
        ws.row_dimensions[r0].height = 48
        for i, rec in enumerate(df.itertuples(index=False), r0 + 1):
            for j, (c, v) in enumerate(zip(cols, rec), 1):
                if isinstance(v, float) and not np.isfinite(v):
                    v = None
                if isinstance(v, pd.Timestamp):
                    v = v.date()
                x = ws.cell(i, j, v)
                x.font, x.border = body_f, bd
                x.alignment = left if isinstance(v, str) else center
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    x.number_format = fmt(c)
                elif hasattr(v, "year"):
                    x.number_format = "dd-mmm-yy"
        ws.freeze_panes = ws.cell(r0 + 1, 3)
        ws.sheet_view.showGridLines = False
        return r0 + len(df)

    wb = Workbook()
    ws = wb.active
    ws.title = "COMPARISON"
    sub = (f"{len(days)} scan days {days[0]:%d-%b-%Y} to {days[-1]:%d-%b-%Y} | 1st half = before {split:%d-%b-%Y} | "
           f"every variant = today's scanner + ONE change | green = better than BASE, red = worse")
    end = sheet(ws, table, "SCANNER IMPROVEMENTS - 6 month test, one change at a time", sub)
    ws.column_dimensions["A"].width, ws.column_dimensions["B"].width = 13, 40
    cols = list(table.columns)
    for j, c in enumerate(cols, 1):
        if c in ("Variant", "What changes") or "Picks" in c or "Entered" in c or "Taken" in c or "Skipped" in c:
            continue
        L = get_column_letter(j)
        base = f"${L}$5"
        rng = f"{L}6:{L}{end}"
        up, down = (red, green) if "SL Hit" in c else (green, red)
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f"AND(ISNUMBER({L}6),{L}6>{base}+0.0000001)"], fill=up))
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f"AND(ISNUMBER({L}6),{L}6<{base}-0.0000001)"], fill=down))
    for j in range(1, len(cols) + 1):
        ws.cell(5, j).fill = base_fill
    for name, (sw, pr) in trades.items():
        if name not in ("BASE", "COMBINED"):
            continue
        for label, t in (("SWING", sw), ("PRIORITY", pr)):
            w = wb.create_sheet(f"{name} {label}")
            show = [c for c in ["Day", "Stock", "Setup", "Regime", "Entered", "Buy Date", "Buy", "Qty", "T1 Hit", "T2 Hit",
                                "SL Hit", "Open", "Exit Date", "PnL %", "PnL Rs"] if c in t.columns]
            t = t[show].copy()
            if "PnL %" in t:
                t["PnL %"] = t["PnL %"] / 100
            sheet(w, t, f"{name} - {label} trades", f"{len(t)} picks")
    w = wb.create_sheet("NOTES")
    sheet(w, pd.DataFrame({"Notes": notes}), "How this test was done", "")
    w.column_dimensions["A"].width = 160
    for ws_ in wb.worksheets:
        ws_.page_setup.orientation = "landscape"
        ws_.sheet_properties.pageSetUpPr.fitToPage = True
    path = Path(path)
    try:
        wb.save(path)
    except PermissionError:
        path = path.with_name(f"{path.stem}_{datetime.now():%H%M}.xlsx")
        wb.save(path)
    return path


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--months", type=int, default=6)
    ap.add_argument("--universe", choices=list(n.NSE_URLS), default="NSE_LISTED")
    ap.add_argument("--period", default="3y")
    ap.add_argument("--cache", default=None, help="folder to keep prices and candidates between runs")
    ap.add_argument("--combined", default="", help="comma list of variants to also run together, e.g. HIST_FIX,EXT7")
    args = ap.parse_args()

    cache = Path(args.cache) if args.cache else None
    pfile = cache / f"prices_{args.universe}_{datetime.now():%Y%m%d}.pkl" if cache else None
    if pfile and pfile.exists():
        data, nifty, sectors, names = pickle.loads(pfile.read_bytes())
        print(f"prices from {pfile}")
    else:
        data, nifty, sectors, names = b6.load(args.universe, args.period)
        if pfile:
            cache.mkdir(parents=True, exist_ok=True)
            pfile.write_bytes(pickle.dumps((data, nifty, sectors, names)))
    last = nifty.index[-1]
    days = nifty.index[(nifty.index >= last - pd.DateOffset(months=args.months)) & (nifty.index < last)]
    print(f"Scan days: {len(days)} ({days[0]:%d-%b-%Y} to {days[-1]:%d-%b-%Y}) | prices up to {last:%d-%b-%Y}")

    cfile = cache / f"candidates_{args.universe}_{datetime.now():%Y%m%d}_{args.months}m.pkl" if cache else None
    if cfile and cfile.exists():
        cands = pd.read_pickle(cfile)
        print(f"candidates from {cfile}")
    else:
        cands = record(data, nifty, sectors, names, days)
        if cfile:
            cands.to_pickle(cfile)
    print(f"{len(cands)} candidates recorded")

    variants = dict(VARIANTS)
    if args.combined:
        combo = {}
        for v in args.combined.split(","):
            combo.update(VARIANTS[v.strip()])
        variants["COMBINED"] = combo
        DESCRIBE["COMBINED"] = "Together: " + " + ".join(v.strip() for v in args.combined.split(","))
    table, trades, split = run_variants(data, cands, days, variants)

    n0, n1 = float(nifty["Close"].loc[days[0]]), float(nifty["Close"].iloc[-1])
    reg = cands.drop_duplicates("Day")["Regime"].value_counts().to_dict()
    notes = [
        f"Test window: {len(days)} trading days, {days[0]:%d-%b-%Y} to {days[-1]:%d-%b-%Y}, prices up to {last:%d-%b-%Y}. "
        f"NIFTY {n0:,.0f} -> {n1:,.0f} ({(n1 / n0 - 1) * 100:+.1f}%). Regime days: {reg}.",
        "Each day the scanner's own rules ran on prices known up to that close only (no look-ahead), smallcap mode, "
        "universe " + args.universe + ". Every variant changes ONE rule; everything else is today's scanner.",
        "SWING = the top 15 picks of each day (score 60+). PRIORITY = the PRIORITY sheet rules on those picks "
        "(BUY ZONE, no risk flags, Hist Hit 7% >= 25, one per sector, max 5 / 3 on BEAR days).",
        "Trades: entry only on the next session inside the entry zone; Target 1 or the stop, whichever first "
        "(same candle = stop); gaps exit at the open; open trades valued at the last close. Same rules as backtest_picks.py.",
        "Avg P&L % = average per trade entered. Profit Factor = total % gained / total % lost (above 1 = profitable).",
        "Rs 1L account: only PRIORITY picks, Suggested Qty (1% risk, max 20% of capital per stock), bought only "
        "if there is free cash; cash comes back on the exit day. P&L per day = account P&L / trading days in the test.",
        "Not included (cannot be tested on old data): news, fundamentals, earnings dates, live price at 9:15.",
        "1st half / 2nd half: a change that only helps in one half is probably luck; keep changes that help in both.",
        "Brokerage, taxes and slippage are not deducted.",
    ]
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    path = write(table, trades, days, split, OUT_DIR / f"Optimization_6Months_{bp.day_tag(datetime.now())}.xlsx", notes)
    print("\nSaved:", path)


if __name__ == "__main__":
    main()
