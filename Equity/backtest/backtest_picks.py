"""Backtest the picks from past swing reports (SwingStocks_* / NIFTY_Swing_*) against real prices.

For every pick this follows the report's own plan on real daily prices: was the entry zone
traded on the purchase date, and then which came first - Target 1 or the Stop Loss.

Install : pip install yfinance pandas numpy openpyxl
Run     : python backtest_picks.py                       (every report in ../output/daily)
          python backtest_picks.py file1.xlsx file2.xlsx (only these reports)
          python backtest_picks.py --all-actions         (also check WAIT / REVIEW rows)
Output  : output/PicksBacktest_<Sun_11th Oct>.xlsx       (named for the day the check was run)
Columns : Sl.No | Stock | Purchase Date | CMP | Entry | Target | Entry Hit? | Target Hit? |
          Gain/Loss % | Gain/Loss Rs | Still Open? | then everything else (result, report name ...)

Rules used (also written into the NOTES sheet of the output)
  * Purchase date = first market session after the report time (the day the report is for).
  * Entry hit = on the purchase date the price traded inside the report's entry zone.
    Buy price = the open if it opened inside the zone; the top of the zone if it opened above
    and came down into it; the bottom of the zone if it opened below and rose into it.
    Never inside the zone that day, or opened at/below the stop = ENTRY NOT HIT, no trade.
  * After buying, the daily candles are walked: whichever of Stop Loss or Target 1 is touched
    first closes the trade. Both on the same day = the stop is counted (conservative).
    If a day opens beyond the level (gap), the exit is at that open.
  * Gain/Loss = exit at Target 1 or Stop Loss; if neither yet, the trade is still OPEN and is
    valued at CMP. Qty = the report's Suggested Qty* (Rs 1L capital, 1% risk); if that was 0
    (very expensive share), 1 share is used and the row is flagged.
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
ENTRY_VALID_DAYS = 1             # the report is for one day: the entry must happen on the purchase date


def day_tag(d):
    """Sun_11th Oct  (same style as the report file names)"""
    suffix = "th" if 11 <= d.day <= 13 else {1: "st", 2: "nd", 3: "rd"}.get(d.day % 10, "th")
    return f"{d:%a}_{d.day}{suffix} {d:%b}"


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
                "Setup": r.get("Setup"), "Swing Score": r.get("Swing Score"),
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


def simulate(bars, lo, hi, stop, t1, t2):
    """Follow the report's plan on daily candles that start on the purchase date.
    Returns entry (bar number or None), buy price, result, exit (bar number or None), exit price, t2 reached."""
    o, h, l = (bars[c].to_numpy(float) for c in ("Open", "High", "Low"))
    res = {"entry": None, "buy": np.nan, "result": "ENTRY NOT HIT", "exit": None, "exit_px": np.nan, "t2": False}
    k, buy, from_above = None, np.nan, False
    for i in range(min(ENTRY_VALID_DAYS, len(o))):
        if o[i] <= stop:                           # opened at/below the stop: the plan is void
            break
        if lo <= o[i] <= hi:
            k, buy = i, o[i]
        elif o[i] > hi and l[i] <= hi:             # opened above the zone, came down into it
            k, buy, from_above = i, hi, True
        elif o[i] < lo and h[i] >= lo:             # opened below the zone, rose into it / trigger hit
            k, buy = i, lo
        if k is not None:
            break
    if k is None:
        return res
    res.update(entry=k, buy=buy, result="OPEN")
    for i in range(k, len(o)):
        first = i == k
        if l[i] <= stop:
            res.update(result="STOP LOSS HIT", exit=i, exit_px=stop if first else min(stop, o[i]))
            return res
        # on the buy day a high made before the dip into the zone must not count as the target
        if h[i] >= t1 and not (first and from_above):
            res.update(result="TARGET HIT", exit=i, exit_px=t1 if first else max(t1, o[i]))
            for j in range(i, len(o)):             # did it go on to Target 2 before falling to the stop?
                if j > i and l[j] <= stop:
                    break
                if h[j] >= t2:
                    res["t2"] = True
                    break
            return res
    return res


def evaluate(row, bars, cmp_price, cmp_src=""):
    """row = one pick with the report's levels. bars = daily candles from the purchase date onwards (or None)."""
    out = {
        "Stock": row["Stock"], "Purchase Date": None, "CMP": cmp_price,
        "Entry Low": row["Entry Low"], "Entry High": row["Entry High"],
        "Target 1": row["Target 1"], "Target 2": row["Target 2"],
        "Entry Hit?": "-", "Target Hit?": "-", "Gain/Loss %": np.nan, "Gain/Loss Rs": np.nan, "Still Open?": "-",
        "Result": "", "Buy Price": np.nan, "Stop Loss": row["Stop Loss"], "Exit Price": np.nan, "Exit Date": None,
        "Days Held": np.nan, "Qty": np.nan, "Invested Rs": np.nan, "Max Gain %": np.nan, "Max Fall %": np.nan,
        "Setup": row.get("Setup"), "Swing Score": row.get("Swing Score"),
        "Report": row.get("Report"), "Generated": row.get("Generated"), "CMP Source": cmp_src, "Note": "",
    }
    if bars is None or bars.empty:
        out["Result"] = "NOT STARTED (no market session after the report yet)"
        return out
    out["Purchase Date"] = bars.index[0].date()
    s = simulate(bars, row["Entry Low"], row["Entry High"], row["Stop Loss"], row["Target 1"], row["Target 2"])
    if s["entry"] is None:
        out.update({"Entry Hit?": "NO", "Result": "ENTRY NOT HIT"})
        return out
    qty = int(row["Qty (report)"]) if pd.notna(row.get("Qty (report)")) else 0
    if qty <= 0:
        qty, out["Note"] = 1, "Report qty was 0 (share too costly for Rs 1L at 1% risk); 1 share used"
    buy, closed = s["buy"], s["exit"] is not None
    held = bars.iloc[s["entry"]:(s["exit"] + 1 if closed else len(bars))]
    exit_px = s["exit_px"] if closed else cmp_price
    out.update({
        "Entry Hit?": "YES", "Result": s["result"], "Still Open?": "NO" if closed else "YES",
        "Target Hit?": ("TARGET 1 & 2" if s["t2"] else "TARGET 1") if s["result"] == "TARGET HIT" else "NO",
        "Buy Price": buy, "Qty": qty, "Invested Rs": buy * qty, "Days Held": len(held),
        "Exit Price": exit_px if closed else np.nan, "Exit Date": held.index[-1].date() if closed else None,
        "Max Gain %": (float(held["High"].max()) / buy - 1) * 100,
        "Max Fall %": (float(held["Low"].min()) / buy - 1) * 100,
    })
    if exit_px is not None and np.isfinite(exit_px):
        out["Gain/Loss %"] = (exit_px / buy - 1) * 100
        out["Gain/Loss Rs"] = (exit_px - buy) * qty
    return out


# ----------------------------------------------------------------------------
# Excel
# ----------------------------------------------------------------------------
MAIN = ["Sl.No", "Stock", "Purchase Date", "CMP", "Entry Low", "Entry High", "Target 1", "Target 2",
        "Entry Hit?", "Target Hit?", "Gain/Loss %", "Gain/Loss Rs", "Still Open?",
        "Result", "Buy Price", "Stop Loss", "Exit Price", "Exit Date", "Days Held", "Qty", "Invested Rs",
        "Max Gain %", "Max Fall %", "Setup", "Swing Score", "NIFTY Regime", "Report", "Generated",
        "CMP Source", "Note"]
COUNTS = {"Sl.No", "Qty", "Days Held", "Swing Score", "Picks", "Entry Hit", "Entry Not Hit", "Target Hit",
          "Stop Loss Hit", "Still Open"}
WIDE = {"Note": 44, "Report": 34, "Best": 22, "Worst": 22, "Result": 30, "Setup": 20, "Generated": 16,
        "Purchase Date": 13, "Exit Date": 13, "Target Hit?": 14, "CMP Source": 13, "Sl.No": 6}
NOTES = [
    "Purchase Date = first market session after the report time (the day the report is for).",
    "Entry Low / Entry High / Target 1 / Target 2 / Stop Loss = the levels the report gave.",
    "Entry Hit? = YES if on the purchase date the price traded inside the entry zone. Buy Price = the open if it "
    "opened inside the zone, the top of the zone if it opened above and came down into it, the bottom of the zone "
    "if it opened below and rose into it. NO = never inside the zone that day (or opened at/below the stop): no trade.",
    "Target Hit? = TARGET 1 if Target 1 was touched before the Stop Loss; TARGET 1 & 2 if it then also reached "
    "Target 2 before falling back to the stop. If both stop and target are touched on the same day the stop is counted.",
    "Gain/Loss % and Rs = exit at Target 1 or at the Stop Loss (at the day's open if it gapped past the level). "
    "If neither is hit yet the trade is Still Open and is valued at CMP.",
    "CMP = live price if the market is open, else the latest close (see CMP Source).",
    "Qty = the report's suggested quantity (Rs 1L capital, 1% risk per trade).",
    "Prices are Yahoo Finance daily candles adjusted for dividends/splits, same as the scanner. "
    "Brokerage, taxes and slippage are not deducted.",
]


def summary(df, by, label):
    """One row per group (report / month / setup ...) plus an ALL row."""
    rows = []
    for name, g in [(k, g) for k, g in df.groupby(by, sort=False)] + [("ALL", df)]:
        ent = g[g["Entry Hit?"] == "YES"]
        tgt, sl = int((ent["Result"] == "TARGET HIT").sum()), int((ent["Result"] == "STOP LOSS HIT").sum())
        row = {
            label: name, "Picks": len(g), "Entry Hit": len(ent), "Entry Not Hit": len(g) - len(ent),
            "Target Hit": tgt, "Stop Loss Hit": sl, "Still Open": int((ent["Result"] == "OPEN").sum()),
            "Win % (closed trades)": tgt / (tgt + sl) * 100 if tgt + sl else np.nan,
            "Avg Gain/Loss %": ent["Gain/Loss %"].mean(), "Total Bought Rs": ent["Invested Rs"].sum(),
            "Total Gain/Loss Rs": ent["Gain/Loss Rs"].sum(),
        }
        v = ent[ent["Gain/Loss %"].notna()]
        if len(v):
            row["Best"] = f"{v.loc[v['Gain/Loss %'].idxmax(), 'Stock']} ({v['Gain/Loss %'].max():+.1f}%)"
            row["Worst"] = f"{v.loc[v['Gain/Loss %'].idxmin(), 'Stock']} ({v['Gain/Loss %'].min():+.1f}%)"
        rows.append(row)
    return pd.DataFrame(rows)


def write(df, path, summaries, notes=NOTES):
    """summaries = {sheet name: summary table}. Returns the path really written."""
    from openpyxl.formatting.rule import CellIsRule, FormulaRule
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    green, red = PatternFill("solid", fgColor="C6EFCE"), PatternFill("solid", fgColor="FFC7CE")
    amber, grey = PatternFill("solid", fgColor="FFEB9C"), PatternFill("solid", fgColor="D9D9D9")
    head = PatternFill("solid", fgColor="1F4E78")

    def style(ws, cols):
        for c in ws[1]:
            c.fill, c.font = head, Font(bold=True, color="FFFFFF", size=10)
            c.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 42
        ws.freeze_panes = "C2" if cols[0] == "Sl.No" else "B2"
        ws.auto_filter.ref = ws.dimensions
        last = ws.max_row
        for j, col in enumerate(cols, 1):
            L = get_column_letter(j)
            ws.column_dimensions[L].width = WIDE.get(col, 12)
            fmt = ("dd-mmm-yy" if "Date" in col else "dd-mmm-yy hh:mm" if col == "Generated"
                   else "0" if col in COUNTS else "0.0" if col.startswith("Win")
                   else "+0.0;-0.0;0.0" if "%" in col
                   else "#,##0;-#,##0" if "Rs" in col else "0.00")
            for i in range(2, last + 1):
                c = ws.cell(i, j)
                if c.value is not None and not isinstance(c.value, str):
                    c.number_format = fmt
            rng = f"{L}2:{L}{last}"
            if "Gain" in col or "Max" in col:
                ws.conditional_formatting.add(rng, CellIsRule(operator="greaterThan", formula=["0"], fill=green))
                ws.conditional_formatting.add(rng, CellIsRule(operator="lessThan", formula=["0"], fill=red))
            if col in ("Entry Hit?", "Target Hit?", "Still Open?", "Result"):
                ws.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT({L}2,6)="TARGET"'], fill=green))
                ws.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT({L}2,4)="STOP"'], fill=red))
                ws.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT({L}2,5)="ENTRY"'], fill=grey))
                ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{L}2="OPEN"'], fill=amber))
                if col == "Entry Hit?":
                    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{L}2="YES"'], fill=green))
                    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{L}2="NO"'], fill=grey))
                if col == "Still Open?":
                    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{L}2="YES"'], fill=amber))

    df = df.copy()
    df.insert(0, "Sl.No", range(1, len(df) + 1))
    cols = [c for c in MAIN if c in df.columns]

    def save(p):
        with pd.ExcelWriter(p, engine="openpyxl") as w:
            df[cols].to_excel(w, sheet_name="PICKS RESULT", index=False)
            style(w.sheets["PICKS RESULT"], cols)
            for name, sm in summaries.items():
                sm.to_excel(w, sheet_name=name, index=False)
                style(w.sheets[name], list(sm.columns))
                w.sheets[name].column_dimensions["A"].width = 36
            pd.DataFrame({"How to read": notes}).to_excel(w, sheet_name="NOTES", index=False)
            w.sheets["NOTES"].column_dimensions["A"].width = 150
            for c in w.sheets["NOTES"]["A"]:
                c.alignment = Alignment(wrap_text=True, vertical="top")
        return p

    path = Path(path)
    try:
        return save(path)
    except PermissionError:          # today's file is open in Excel
        return save(path.with_name(f"{path.stem}_{datetime.now():%H%M}.xlsx"))


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
    picks = picks.sort_values("Generated", kind="stable")
    syms = sorted(picks["Stock"].unique())
    hist = fetch_history(syms, picks["Generated"].min() - timedelta(days=7))
    print(f"price history for {len(hist)}/{len(syms)} stocks")
    cmp = fetch_cmp(syms, hist)
    rows = []
    for r in picks.to_dict("records"):
        d, (price, src) = hist.get(r["Stock"]), cmp.get(r["Stock"], (np.nan, ""))
        if d is None:
            rows.append({**evaluate(r, None, price, src), "Result": "NO PRICE DATA"})
            continue
        td = trade_day(d.index, r["Generated"])
        rows.append(evaluate(r, d.loc[td:] if td is not None else None, price, src))
    res = pd.DataFrame(rows)
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    path = write(res, OUT_DIR / f"PicksBacktest_{day_tag(datetime.now())}.xlsx",
                 {"SUMMARY": summary(res, "Report", "Report")})
    show = ["Stock", "Purchase Date", "CMP", "Entry Low", "Entry High", "Target 1", "Entry Hit?", "Target Hit?",
            "Gain/Loss %", "Gain/Loss Rs", "Still Open?"]
    print(res[show].to_string(index=False, float_format=lambda v: f"{v:,.2f}"))
    print("\nSaved:", path)


if __name__ == "__main__":
    main()
