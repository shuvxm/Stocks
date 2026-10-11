"""Backtest the BUY ZONE picks from past swing reports (SwingStocks_* / NIFTY_Swing_*) against real prices.

For every SWING PICKS row whose Action was "BUY ZONE", this checks what the stock
did after the report: price now (CMP), gain/loss in % and Rs, and whether the
report's stop loss or targets were hit first.

Install : pip install yfinance pandas numpy openpyxl
Run     : python backtest_picks.py                       (every report in ../output/daily)
          python backtest_picks.py file1.xlsx file2.xlsx (only these reports)
          python backtest_picks.py --all-actions         (also check WAIT rows)
Output  : output/Picks_Backtest_<date_time>.xlsx

Rules used (also written into the NOTES sheet of the output)
  * Buy price = the "Live Price" the report showed (the price it said was in the buy zone).
    The next session's open is shown too, because a report made after 3:30 PM or before
    9:15 AM can only really be bought at the next open.
  * Trade day = first market session after the report time.
  * Qty = the report's Suggested Qty* (Rs 1L capital, 1% risk). If that was 0 (very
    expensive share), 1 share is used and the row is flagged.
  * Plan result walks the daily candles from the trade day: whichever of Stop Loss or
    Target 1 is touched first decides it. If both are touched on the same day the stop is
    assumed first (conservative). Nothing touched yet = still OPEN, valued at CMP.
"""
import argparse
import re
import sys
from datetime import datetime, timedelta
from pathlib import Path

import numpy as np
import pandas as pd

BASE = Path(__file__).resolve().parent
REPORT_DIR = BASE.parent / "output" / "daily"       # the screener saves its reports in Equity/output/daily
REPORT_GLOBS = ("NIFTY_Swing_*.xlsx", "SwingStocks_*.xlsx")   # old and new report names
OUT_DIR = BASE / "output"
MARKET_OPEN = (9, 15)


# ----------------------------------------------------------------------------
# Read reports
# ----------------------------------------------------------------------------
def report_time(path):
    """Generated time from the subtitle row, else from the file name."""
    try:
        sub = str(pd.read_excel(path, "SWING PICKS", header=None, nrows=2).iloc[1, 0])
        m = re.search(r"Generated:\s*(\d{2}-\w{3}-\d{4} \d{2}:\d{2})", sub)
        if m:
            return datetime.strptime(m.group(1), "%d-%b-%Y %H:%M")
    except Exception:
        pass
    m = re.search(r"(\d{4}-\d{2}-\d{2})(?:_(\d{4}))?", Path(path).name)
    if m:
        return datetime.strptime(m.group(1) + (m.group(2) or "1600"), "%Y-%m-%d%H%M")
    raise ValueError(f"cannot find report date in {path}")


def load_picks(paths, all_actions):
    rows = []
    for p in paths:
        gen = report_time(p)
        df = pd.read_excel(p, "SWING PICKS", header=3)
        df = df[df["Stock"].notna()]
        if not all_actions:
            df = df[df["Action"].astype(str).str.startswith("BUY ZONE")]
        for _, r in df.iterrows():
            rows.append({
                "Report": Path(p).name, "Generated": gen, "Stock": str(r["Stock"]).strip().upper(),
                "Setup": r.get("Setup"), "Action (report)": r.get("Action"),
                "Swing Score": r.get("Swing Score"), "Grade": r.get("Grade"),
                "Report Price": r.get("Live Price") if pd.notna(r.get("Live Price")) else r.get("Last Close"),
                "Entry Low": r.get("Entry Low"), "Entry High": r.get("Entry High"),
                "Stop Loss": r.get("Stop Loss"), "Target 1": r.get("Target 1"), "Target 2": r.get("Target 2"),
                "Qty (report)": r.get("Suggested Qty*"),
            })
    return pd.DataFrame(rows)


# ----------------------------------------------------------------------------
# Prices
# ----------------------------------------------------------------------------
def fetch_history(symbols, start):
    import yfinance as yf
    tickers = [s + ".NS" for s in symbols]
    raw = yf.download(tickers, start=start.strftime("%Y-%m-%d"), interval="1d", auto_adjust=True,
                      progress=False, threads=True, group_by="ticker")
    out = {}
    for s, t in zip(symbols, tickers):
        try:
            d = raw[t] if isinstance(raw.columns, pd.MultiIndex) else raw
            d = d[["Open", "High", "Low", "Close"]].apply(pd.to_numeric, errors="coerce").dropna()
            if len(d):
                d.index = pd.to_datetime(d.index).tz_localize(None)
                out[s] = d
        except Exception:
            pass
    return out


def fetch_cmp(symbols, hist):
    """Live price if the market is open, else the last close."""
    import yfinance as yf
    cmp = {}
    for s in symbols:
        price, src = np.nan, ""
        try:
            q = float(yf.Ticker(s + ".NS").fast_info["last_price"])
            if q > 0 and np.isfinite(q):
                price, src = q, "Live quote"
        except Exception:
            pass
        if not np.isfinite(price) and s in hist:
            price, src = float(hist[s]["Close"].iloc[-1]), f"Close {hist[s].index[-1]:%d-%b}"
        cmp[s] = (price, src)
    return cmp


# ----------------------------------------------------------------------------
# Evaluate one pick
# ----------------------------------------------------------------------------
def trade_day(index, gen):
    """First session that opens after the report was generated."""
    open_at = gen.replace(hour=MARKET_OPEN[0], minute=MARKET_OPEN[1], second=0, microsecond=0)
    first = gen.date() if gen < open_at else gen.date() + timedelta(days=1)
    later = index[index.date >= first]
    return later[0] if len(later) else None


def plan_result(bars, stop, t1, t2):
    """Walk candles: first of SL / T1 decides. Returns (result, date, exit price, t2 hit)."""
    for dt, b in bars.iterrows():
        hit_sl, hit_t1 = b["Low"] <= stop, b["High"] >= t1
        if hit_sl:
            return "STOP LOSS HIT", dt, stop, False
        if hit_t1:
            after = bars.loc[dt:]
            t2_hit = bool((after["High"] >= t2).any())
            return ("TARGET 2 HIT" if t2_hit else "TARGET 1 HIT"), dt, (t2 if t2_hit else t1), t2_hit
    return "OPEN", None, None, False


def evaluate(row, hist, cmp):
    s = row["Stock"]
    out = dict(row)
    d = hist.get(s)
    price, src = cmp.get(s, (np.nan, ""))
    out["CMP"], out["CMP Source"] = price, src
    if d is None or d.empty:
        out["Plan Result"] = "NO PRICE DATA"
        return out
    td = trade_day(d.index, row["Generated"])
    if td is None:
        out["Plan Result"] = "NOT STARTED (no session after report yet)"
        return out
    bars = d.loc[td:]
    buy = float(row["Report Price"])
    qty = int(row["Qty (report)"]) if pd.notna(row["Qty (report)"]) else 0
    flag = ""
    if qty <= 0:
        qty, flag = 1, "Report qty was 0 (share too costly for Rs 1L at 1% risk); 1 share used"
    nxt_open = float(bars["Open"].iloc[0])

    out.update({
        "Trade Day": td.date(), "Buy Price": buy, "Next Day Open": nxt_open,
        "Gap at Open %": (nxt_open / buy - 1) * 100,
        "Opened In Entry Zone?": "Yes" if row["Entry Low"] * 0.995 <= nxt_open <= row["Entry High"] * 1.005 else "No",
        "Qty Used": qty, "Investment Rs": buy * qty,
        "Days Held": len(bars),
        "Highest Since": float(bars["High"].max()), "Lowest Since": float(bars["Low"].min()),
    })
    out["Max Gain %"] = (out["Highest Since"] / buy - 1) * 100
    out["Max Fall %"] = (out["Lowest Since"] / buy - 1) * 100
    if np.isfinite(price):
        out["Gain/Loss %"] = (price / buy - 1) * 100
        out["Gain/Loss Rs"] = (price - buy) * qty
        out["Gain/Loss % (bought at next open)"] = (price / nxt_open - 1) * 100
        out["Gain/Loss Rs (bought at next open)"] = (price - nxt_open) * qty
    res, when, exit_px, _ = plan_result(bars, row["Stop Loss"], row["Target 1"], row["Target 2"])
    out["Plan Result"], out["Plan Exit Date"] = res, (when.date() if when is not None else "")
    exit_px = exit_px if exit_px is not None else price
    out["Plan Exit Price"] = exit_px
    if exit_px is not None and np.isfinite(exit_px):
        out["Plan P/L %"] = (exit_px / buy - 1) * 100
        out["Plan P/L Rs"] = (exit_px - buy) * qty
    out["Status Now"] = ("UP" if out.get("Gain/Loss %", 0) > 0 else "DOWN") if np.isfinite(price) else ""
    out["Note"] = flag
    return out


# ----------------------------------------------------------------------------
# Excel
# ----------------------------------------------------------------------------
MAIN = ["Stock", "Report", "Generated", "Trade Day", "Setup", "Swing Score", "Entry Low", "Entry High",
        "Buy Price", "Next Day Open", "Gap at Open %", "Opened In Entry Zone?", "CMP", "CMP Source",
        "Status Now", "Gain/Loss %", "Qty Used", "Investment Rs", "Gain/Loss Rs",
        "Gain/Loss % (bought at next open)", "Gain/Loss Rs (bought at next open)",
        "Stop Loss", "Target 1", "Target 2", "Highest Since", "Lowest Since", "Max Gain %", "Max Fall %",
        "Days Held", "Plan Result", "Plan Exit Date", "Plan Exit Price", "Plan P/L %", "Plan P/L Rs", "Note"]
PCT = {c for c in MAIN if "%" in c}
RS = {"Investment Rs", "Gain/Loss Rs", "Gain/Loss Rs (bought at next open)", "Plan P/L Rs"}


def summary(df):
    rows = []
    for name, g in list(df.groupby("Report", sort=False)) + [("ALL REPORTS", df)]:
        g = g[g["Gain/Loss %"].notna()] if "Gain/Loss %" in g else g.iloc[0:0]
        if g.empty:
            continue
        rows.append({
            "Report": name, "Buy Zone Picks": len(g),
            "Up Now": int((g["Gain/Loss %"] > 0).sum()), "Down Now": int((g["Gain/Loss %"] <= 0).sum()),
            "Avg Gain/Loss %": g["Gain/Loss %"].mean(), "Total Invested Rs": g["Investment Rs"].sum(),
            "Total Gain/Loss Rs (hold till now)": g["Gain/Loss Rs"].sum(),
            "Hit Target 1+": int(g["Plan Result"].str.startswith("TARGET").sum()),
            "Hit Stop Loss": int((g["Plan Result"] == "STOP LOSS HIT").sum()),
            "Still Open": int((g["Plan Result"] == "OPEN").sum()),
            "Total Plan P/L Rs (follow SL/T1)": g["Plan P/L Rs"].sum(),
            "Best": f"{g.loc[g['Gain/Loss %'].idxmax(), 'Stock']} ({g['Gain/Loss %'].max():+.1f}%)",
            "Worst": f"{g.loc[g['Gain/Loss %'].idxmin(), 'Stock']} ({g['Gain/Loss %'].min():+.1f}%)",
        })
    return pd.DataFrame(rows)


def write(df, path):
    from openpyxl.formatting.rule import CellIsRule, FormulaRule
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    green, red = PatternFill("solid", fgColor="C6EFCE"), PatternFill("solid", fgColor="FFC7CE")
    amber = PatternFill("solid", fgColor="FFEB9C")
    head = PatternFill("solid", fgColor="1F4E78")

    def style(ws, cols):
        for c in ws[1]:
            c.fill, c.font = head, Font(bold=True, color="FFFFFF", size=10)
            c.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 42
        ws.freeze_panes = "B2"
        ws.auto_filter.ref = ws.dimensions
        last = ws.max_row
        for j, col in enumerate(cols, 1):
            L = get_column_letter(j)
            ws.column_dimensions[L].width = 40 if col in ("Note", "Report") else 16 if col in ("Plan Result", "Generated") else 12
            fmt = "+0.0;-0.0;0.0" if col in PCT else "#,##0;-#,##0" if col in RS or "Rs" in col else "0.00"
            for i in range(2, last + 1):
                c = ws.cell(i, j)
                if isinstance(c.value, (int, float)) and col not in ("Qty Used", "Days Held", "Swing Score",
                                                                      "Buy Zone Picks", "Up Now", "Down Now",
                                                                      "Hit Target 1+", "Hit Stop Loss", "Still Open"):
                    c.number_format = fmt
                if col == "Generated" and c.value is not None:
                    c.number_format = "dd-mmm hh:mm"
            if col in PCT or "Rs" in col and col not in ("Investment Rs", "Total Invested Rs"):
                rng = f"{L}2:{L}{last}"
                ws.conditional_formatting.add(rng, CellIsRule(operator="greaterThan", formula=["0"], fill=green))
                ws.conditional_formatting.add(rng, CellIsRule(operator="lessThan", formula=["0"], fill=red))
            if col in ("Plan Result", "Status Now"):
                rng = f"{L}2:{L}{last}"
                ws.conditional_formatting.add(rng, FormulaRule(formula=[f'OR(LEFT({L}2,6)="TARGET",{L}2="UP")'], fill=green))
                ws.conditional_formatting.add(rng, FormulaRule(formula=[f'OR(LEFT({L}2,4)="STOP",{L}2="DOWN")'], fill=red))
                ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{L}2="OPEN"'], fill=amber))

    cols = [c for c in MAIN if c in df.columns]
    sm = summary(df)
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        sm.to_excel(w, sheet_name="SUMMARY", index=False)
        style(w.sheets["SUMMARY"], list(sm.columns))
        w.sheets["SUMMARY"].column_dimensions["A"].width = 38
        df[cols].to_excel(w, sheet_name="PICKS RESULT", index=False)
        style(w.sheets["PICKS RESULT"], cols)
        notes = pd.DataFrame({"How to read": [
            "Buy Price = the price the report showed when it said BUY ZONE.",
            "Next Day Open = what you could really have paid if you bought at the next market open. "
            "Gap at Open % shows how far that was from the report price.",
            "CMP = live price if the market is open, else the latest close (see CMP Source).",
            "Gain/Loss % and Rs = holding from Buy Price until now, with the report's suggested quantity.",
            "Plan Result = following the report's levels: whichever of Stop Loss or Target 1 was touched first. "
            "If both on the same day, the stop is counted (conservative). OPEN = neither touched yet.",
            "Plan P/L = exit at the stop or target that was hit, or CMP if still open.",
            "Prices are Yahoo Finance daily candles adjusted for dividends/splits, same as the scanner.",
            "A few days of results is a small sample: use it to spot problems, not to prove the strategy.",
        ]})
        notes.to_excel(w, sheet_name="NOTES", index=False)
        w.sheets["NOTES"].column_dimensions["A"].width = 140
    return path


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("reports", nargs="*", help="report files (default: every report in ../output/daily)")
    ap.add_argument("--all-actions", action="store_true", help="also check WAIT / REVIEW rows")
    args = ap.parse_args()
    paths = [Path(p) for p in args.reports] or [p for g in REPORT_GLOBS for p in sorted(REPORT_DIR.glob(g))]
    paths = [p for p in paths if not p.name.startswith("~$")]
    if not paths:
        sys.exit(f"No swing reports found in {REPORT_DIR}")
    picks = load_picks(paths, args.all_actions)
    print(f"{len(picks)} picks from {len(paths)} report(s)")
    if picks.empty:
        sys.exit("No BUY ZONE picks in these reports.")
    syms = sorted(picks["Stock"].unique())
    hist = fetch_history(syms, picks["Generated"].min() - timedelta(days=7))
    print(f"price history for {len(hist)}/{len(syms)} stocks")
    cmp = fetch_cmp(syms, hist)
    res = pd.DataFrame([evaluate(r, hist, cmp) for r in picks.to_dict("records")])
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    path = write(res, OUT_DIR / f"Picks_Backtest_{datetime.now():%Y-%m-%d_%H%M}.xlsx")
    show = [c for c in ["Stock", "Trade Day", "Buy Price", "CMP", "Gain/Loss %", "Gain/Loss Rs", "Plan Result"] if c in res]
    print(res[show].to_string(index=False, float_format=lambda v: f"{v:,.2f}"))
    print("\nSaved:", path)


if __name__ == "__main__":
    main()
