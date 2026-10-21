"""Backtest the PRIORITY and SWING PICKS sheets of the daily SwingStocks_* reports against real prices.

For every stock on either sheet this follows the report's own plan on real daily candles:
was the entry zone (Entry Low - Entry High) traded on the day the report is for, and after buying,
did the price reach Target 1 / Target 2 before the Stop Loss.

Install : pip install yfinance pandas numpy openpyxl
Run     : python backtest_report_sheets.py                        (every SwingStocks_* report in ../output/daily)
          python backtest_report_sheets.py file1.xlsx file2.xlsx  (only these reports)
Output  : ../../REPORTS/Backtest_Priority_SwingPicks_<first>_to_<last>.xlsx

Rules (also written into the RULES sheet of the output)
  * Entry date = first market session after the report was generated (the day the report is for).
  * Entry hit = on the entry date the price traded inside the entry zone. Buy price = the open if it
    opened inside the zone; Entry High if it opened above and came down into it; Entry Low if it opened
    below and rose into it. Opened at/below the stop, or never inside the zone that day = no trade.
  * After buying the daily candles are walked to the latest candle:
      Target 1 hit  = Target 1 touched before the Stop Loss.
      Target 2 hit  = Target 2 touched before the Stop Loss.
      Stop Loss hit = Stop Loss touched before Target 1 (a losing trade).
    Stop and target on the same candle = the stop is counted (we cannot know which came first).
    On the entry day, if the stock opened above the zone and dipped into it, that day's high is not
    counted as a target (it was most likely made before the dip).
    A candle that opens beyond a level (gap) exits at that open.
  * P&L "book at T1" = sell everything at Target 1 (or the stop). P&L "hold for T2" = keep the full
    position until Target 2 or the stop. Trades with neither yet are valued at the latest price.
"""
import argparse
import re
import sys
from datetime import datetime, time
from pathlib import Path

import numpy as np
import pandas as pd

from backtest_picks import REPORT_DIR, fetch_history, report_time, trade_day

BASE = Path(__file__).resolve().parent
OUT_DIR = BASE.parents[1] / "REPORTS"
SHEETS = ("PRIORITY", "SWING PICKS")


# ----------------------------------------------------------------------------
# Read reports
# ----------------------------------------------------------------------------
def report_day(path):
    """'Tue_13th Oct' from SwingStocks_Tue_13th Oct.xlsx"""
    m = re.search(r"SwingStocks_(.+?)\.xlsx$", Path(path).name)
    return m.group(1).replace("_", " ") if m else Path(path).stem


def read_sheet(path, sheet):
    """Rows of the main table of a sheet. PRIORITY: header row starts with 'Priority' and the table
    ends at 'OTHER PICKS LEFT OUT'. SWING PICKS: header row starts with 'Rank'."""
    raw = pd.read_excel(path, sheet, header=None)
    key = "Priority" if sheet == "PRIORITY" else "Rank"
    hi = raw.index[raw[0].astype(str).str.strip() == key][0]
    df = raw.iloc[hi + 1:].copy()
    df.columns = [str(c).strip() for c in raw.iloc[hi]]
    stop = df.index[df[key].astype(str).str.startswith("OTHER PICKS")]
    if len(stop):
        df = df.loc[:stop[0] - 1]
    df = df[pd.to_numeric(df[key], errors="coerce").notna() & df["Stock"].notna()]
    # PRIORITY also has a 'Rank' column (the stock's SWING PICKS rank): keep it as 'Swing Rank'
    return df.rename(columns={"Rank": "Swing Rank", "Priority": "Rank"}) if sheet == "PRIORITY" else df


def load(paths):
    rows = []
    for p in paths:
        gen, day = report_time(p), report_day(p)
        regime = re.search(r"NIFTY regime:\s*(\w+)", str(pd.read_excel(p, "SWING PICKS", header=None, nrows=2).iloc[1, 0]))
        for sheet in SHEETS:
            for _, r in read_sheet(p, sheet).iterrows():
                rows.append({
                    "Sheet": sheet, "Report": Path(p).name, "Report Day": day, "Generated": gen,
                    "Regime": regime.group(1) if regime else "", "Rank": int(r["Rank"]),
                    "Stock": str(r["Stock"]).strip().upper(), "Setup": r.get("Setup"), "Action": r.get("Action"),
                    "Hist Hit 7%": r.get("Hist Hit 7%"), "Report Price": r.get("Live Price"),
                    "Entry Low": float(r["Entry Low"]), "Entry High": float(r["Entry High"]),
                    "Stop Loss": float(r["Stop Loss"]), "Target 1": float(r["Target 1"]),
                    "Target 2": float(r["Target 2"]), "Qty": r.get("Suggested Qty*"),
                })
    return pd.DataFrame(rows)


# ----------------------------------------------------------------------------
# Follow one trade
# ----------------------------------------------------------------------------
def simulate(bars, lo, hi, stop, t1, t2):
    o, h, l = (bars[c].to_numpy(float) for c in ("Open", "High", "Low"))
    dates = [d.date() for d in bars.index]
    res = {"entry": "NO", "entry_type": "", "buy": np.nan, "t1": None, "t1_px": np.nan, "t2": None,
           "t2_px": np.nan, "sl": None, "sl_px": np.nan, "sl_after_t1": None, "sl_after_t1_px": np.nan,
           "note": [], "max_hi": np.nan, "max_lo": np.nan}
    # entry: the entry date only
    if o[0] <= stop:
        res["entry_type"] = "Opened at/below Stop Loss - skipped"
        return res
    if lo <= o[0] <= hi:
        buy, kind = o[0], "Opened inside zone - bought at open"
    elif o[0] > hi and l[0] <= hi:
        buy, kind = hi, "Opened above zone, dipped in - bought at Entry High"
    elif o[0] < lo and h[0] >= lo:
        buy, kind = lo, "Opened below zone, rose in - bought at Entry Low"
    else:
        res["entry_type"] = "Opened above zone, never came back" if o[0] > hi else "Never rose to Entry Low"
        return res
    res.update(entry="YES", entry_type=kind, buy=buy)
    from_above = o[0] > hi

    last = len(o)
    for i in range(last):
        first = i == 0
        hi_ok = not (first and from_above)        # entry-day high was made before the dip
        stop_hit = l[i] <= stop
        stop_px = stop if first else min(stop, o[i])
        if res["t1"] is None:
            if stop_hit:
                if hi_ok and h[i] >= t1:
                    res["note"].append(f"Stop and Target 1 both touched on {dates[i]:%d-%b}: counted as Stop Loss")
                res.update(sl=dates[i], sl_px=stop_px)
                break
            if hi_ok and h[i] >= t1:
                res.update(t1=dates[i], t1_px=t1 if first else max(t1, o[i]))
                if h[i] >= t2:
                    res.update(t2=dates[i], t2_px=t2 if first else max(t2, o[i]))
                    break
        else:                                     # after Target 1: wait for Target 2 or the stop
            if stop_hit:
                if h[i] >= t2:
                    res["note"].append(f"Stop and Target 2 both touched on {dates[i]:%d-%b}: counted as Stop Loss")
                res.update(sl_after_t1=dates[i], sl_after_t1_px=stop_px)
                break
            if h[i] >= t2:
                res.update(t2=dates[i], t2_px=max(t2, o[i]))
                break
    end = i + 1
    res["max_hi"] = (float(np.max(h[:end][1:] if from_above and end > 1 else h[:end])) / buy - 1) * 100 \
        if not (from_above and end == 1) else np.nan
    res["max_lo"] = (float(np.min(l[:end])) / buy - 1) * 100
    return res


def evaluate(row, bars, cmp_price, cmp_date):
    out = dict(row)
    out.update({"Entry Date": None, "Open on Entry Day": np.nan, "Entry Hit?": "-", "Entry Type": "",
                "Buy Price": np.nan, "T1 Hit?": "-", "T1 Date": None, "T2 Hit?": "-", "T2 Date": None,
                "SL Hit?": "-", "SL Date": None, "Result": "", "Max Gain %": np.nan, "Max Fall %": np.nan,
                "Latest Price": cmp_price, "Latest Price Date": cmp_date,
                "P&L % (book at T1)": np.nan, "P&L Rs (book at T1)": np.nan,
                "P&L % (hold for T2)": np.nan, "P&L Rs (hold for T2)": np.nan, "Note": ""})
    if bars is None or bars.empty:
        out["Result"] = "NO PRICE DATA"
        return out
    out["Entry Date"], out["Open on Entry Day"] = bars.index[0].date(), float(bars["Open"].iloc[0])
    s = simulate(bars, row["Entry Low"], row["Entry High"], row["Stop Loss"], row["Target 1"], row["Target 2"])
    out["Entry Type"] = s["entry_type"]
    if s["entry"] == "NO":
        out.update({"Entry Hit?": "NO", "Result": "ENTRY NOT HIT"})
        return out
    buy = s["buy"]
    qty = int(row["Qty"]) if pd.notna(row["Qty"]) and int(row["Qty"]) > 0 else 1
    notes = s["note"] + ([] if pd.notna(row["Qty"]) and int(row["Qty"]) > 0 else ["Report qty was 0: 1 share used"])
    t1, t2, sl, sl2 = s["t1"], s["t2"], s["sl"], s["sl_after_t1"]
    if t2:
        result = "TARGET 2 HIT"
    elif t1 and sl2:
        result = "TARGET 1 HIT, then fell to Stop Loss"
    elif t1:
        result = "TARGET 1 HIT (T2 still open)"
    elif sl:
        result = "STOP LOSS HIT"
    else:
        result = "OPEN (no target / stop yet)"
    exit_t1 = s["t1_px"] if t1 else s["sl_px"] if sl else cmp_price
    exit_t2 = s["t2_px"] if t2 else s["sl_after_t1_px"] if sl2 else s["sl_px"] if sl else cmp_price
    out.update({
        "Entry Hit?": "YES", "Buy Price": buy, "Result": result,
        "T1 Hit?": "YES" if t1 else "NO", "T1 Date": t1, "T2 Hit?": "YES" if t2 else "NO", "T2 Date": t2,
        "SL Hit?": "YES" if (sl or sl2) else "NO", "SL Date": sl or sl2,
        "Max Gain %": s["max_hi"], "Max Fall %": s["max_lo"], "Qty": qty,
        "P&L % (book at T1)": (exit_t1 / buy - 1) * 100, "P&L Rs (book at T1)": (exit_t1 - buy) * qty,
        "P&L % (hold for T2)": (exit_t2 / buy - 1) * 100, "P&L Rs (hold for T2)": (exit_t2 - buy) * qty,
        "Note": "; ".join(notes),
    })
    return out


# ----------------------------------------------------------------------------
# Summaries
# ----------------------------------------------------------------------------
def summarise(g, label):
    ent = g[g["Entry Hit?"] == "YES"]
    n = len(ent)
    t1, t2 = int((ent["T1 Hit?"] == "YES").sum()), int((ent["T2 Hit?"] == "YES").sum())
    sl = int((ent["Result"] == "STOP LOSS HIT").sum())
    t1_sl = int((ent["Result"] == "TARGET 1 HIT, then fell to Stop Loss").sum())
    op = int(ent["Result"].str.startswith("OPEN").sum())
    pct = lambda k: k / n if n else np.nan
    return {
        "Group": label, "Picks": len(g), "Entry Hit": n, "Entry Not Hit": len(g) - n,
        "Entry Hit %": n / len(g) if len(g) else np.nan,
        "Target 1 Hit": t1, "T1 Hit %": pct(t1), "Target 2 Hit": t2, "T2 Hit %": pct(t2),
        "Stop Loss Hit (before T1)": sl, "SL Hit %": pct(sl),
        "T1 then SL (before T2)": t1_sl, "Still Open": op, "Open %": pct(op),
        "Win % (closed: T1 vs SL)": t1 / (t1 + sl) if t1 + sl else np.nan,
        "Avg P&L % (book at T1)": ent["P&L % (book at T1)"].mean() / 100 if n else np.nan,
        "Total P&L Rs (book at T1)": ent["P&L Rs (book at T1)"].sum(),
        "Avg P&L % (hold for T2)": ent["P&L % (hold for T2)"].mean() / 100 if n else np.nan,
        "Total P&L Rs (hold for T2)": ent["P&L Rs (hold for T2)"].sum(),
        "Total Bought Rs (sum of trades)": (ent["Buy Price"] * ent["Qty"]).sum(),
    }


def summary_tables(res):
    pr, sw = res[res["Sheet"] == "PRIORITY"], res[res["Sheet"] == "SWING PICKS"]
    pkey = set(zip(pr["Report"], pr["Stock"]))
    sw_rest = sw[[(a, b) not in pkey for a, b in zip(sw["Report"], sw["Stock"])]]
    overall = pd.DataFrame([summarise(pr, "PRIORITY sheet"), summarise(sw, "SWING PICKS sheet (all)"),
                            summarise(sw_rest, "SWING PICKS not on PRIORITY"),
                            summarise(sw[sw["Action"].astype(str).str.startswith("BUY ZONE")], "SWING PICKS - BUY ZONE rows"),
                            summarise(sw[sw["Action"].astype(str).str.startswith("WAIT")], "SWING PICKS - WAIT rows")])
    days = list(dict.fromkeys(res.sort_values("Generated")["Report Day"]))
    by_day = {s: pd.DataFrame([summarise(res[(res["Sheet"] == s) & (res["Report Day"] == d)],
                                         f"{d} ({res.loc[res['Report Day'] == d, 'Regime'].iloc[0]})") for d in days]
                              + [summarise(res[res["Sheet"] == s], "ALL DAYS")]) for s in SHEETS}
    return overall, by_day


# ----------------------------------------------------------------------------
# Excel
# ----------------------------------------------------------------------------
TRADE_COLS = ["Report Day", "Regime", "Rank", "Stock", "Setup", "Action", "Hist Hit 7%", "Report Price",
              "Entry Low", "Entry High", "Stop Loss", "Target 1", "Target 2", "Entry Date", "Open on Entry Day",
              "Entry Hit?", "Entry Type", "Buy Price", "Qty", "T1 Hit?", "T1 Date", "T2 Hit?", "T2 Date",
              "SL Hit?", "SL Date", "Result", "Max Gain %", "Max Fall %", "Latest Price", "Latest Price Date",
              "P&L % (book at T1)", "P&L Rs (book at T1)", "P&L % (hold for T2)", "P&L Rs (hold for T2)", "Note"]
WIDTH = {"Report Day": 13, "Regime": 9, "Rank": 6, "Stock": 13, "Setup": 18, "Action": 24, "Entry Type": 34,
         "Result": 30, "Note": 48, "Group": 30, "Qty": 6}
RULES = [
    ("What was tested", "Every stock on the PRIORITY sheet and every stock on the SWING PICKS sheet of each daily "
     "report, using the report's own Entry Low / Entry High / Stop Loss / Target 1 / Target 2."),
    ("Entry date", "The first market session after the report was generated, i.e. the day in the file name. "
     "20-Oct-2026 was a market holiday, so the 'Wed 21st Oct' report (made on 19-Oct) trades on 21-Oct."),
    ("Entry Hit?", "YES only if on the entry date the price traded inside Entry Low - Entry High. Buy Price = the open "
     "if it opened inside the zone; Entry High if it opened above and came down into it; Entry Low if it opened below "
     "and rose into it. NO = the price never was inside the zone that day, or it opened at/below the Stop Loss."),
    ("Target 1 Hit", "Target 1 (+7% of the report price) was touched before the Stop Loss."),
    ("Target 2 Hit", "Target 2 (+10% of the report price) was touched before the Stop Loss. Every T2 hit is also a T1 hit."),
    ("Stop Loss Hit (before T1)", "The Stop Loss was touched before Target 1: a losing trade."),
    ("T1 then SL (before T2)", "Target 1 was reached, then the price fell back to the Stop Loss before Target 2. "
     "Booking at T1 this is a win; holding for T2 it is a loss."),
    ("Still Open", "Bought, but neither Target 1 nor the Stop Loss has been touched yet."),
    ("Percentages", "T1 Hit %, T2 Hit %, SL Hit % and Open % are out of the trades that were entered (Entry Hit). "
     "Entry Hit % is out of all picks. Win % = T1 hits / (T1 hits + SL hits), i.e. only trades that have finished."),
    ("Same-candle rule", "Only daily candles are available, so the order of the high and low inside one day is unknown. "
     "If the Stop Loss and a target are both touched on the same candle, the Stop Loss is counted (conservative). "
     "On the entry day, if the stock opened above the zone and dipped into it, that day's high is not counted as a target."),
    ("Gaps", "If a later candle opens beyond the stop or a target, the exit is at that open price."),
    ("P&L (book at T1)", "Sell the whole position at Target 1, or at the Stop Loss. Open trades valued at the latest price."),
    ("P&L (hold for T2)", "Keep the whole position until Target 2, or the Stop Loss. Open trades valued at the latest price."),
    ("Quantity", "The report's Suggested Qty* (Rs 1,00,000 capital, 1% risk per trade). Rs P&L = price move x Qty."),
    ("Same stock twice", "A stock on both sheets of one report is one real trade; it is counted once in each sheet's "
     "own table. 'SWING PICKS not on PRIORITY' removes those, to show what the PRIORITY filter left out."),
    ("Prices", "Yahoo Finance daily candles (NSE), adjusted, the same source the scanner uses. Brokerage, taxes and "
     "slippage are not deducted."),
]


def write(res, overall, by_day, path, asof):
    from openpyxl import Workbook
    from openpyxl.formatting.rule import FormulaRule
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    F = "Arial"
    title_f, sub_f = Font(name=F, size=14, bold=True, color="FFFFFF"), Font(name=F, size=9, italic=True, color="404040")
    head_f, body_f, bold_f = Font(name=F, size=9, bold=True, color="FFFFFF"), Font(name=F, size=9), Font(name=F, size=9, bold=True)
    sec_f = Font(name=F, size=11, bold=True, color="1F4E78")
    title_fill, head_fill = PatternFill("solid", fgColor="1F4E78"), PatternFill("solid", fgColor="2E75B6")
    all_fill = PatternFill("solid", fgColor="DDEBF7")
    green, red = PatternFill("solid", fgColor="C6EFCE"), PatternFill("solid", fgColor="FFC7CE")
    amber, grey = PatternFill("solid", fgColor="FFEB9C"), PatternFill("solid", fgColor="E7E6E6")
    dgreen = PatternFill("solid", fgColor="63BE7B")
    thin = Side(style="thin", color="B4C6E7")
    bd = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)

    def fmt(col):
        if "Date" in col:
            return "dd-mmm-yy"
        if col.endswith("Hit %") or col in ("Open %", "Entry Hit %") or col.startswith(("Win %", "Avg P&L")):
            return "0.0%"
        if col.startswith("P&L %") or col.startswith("Max") or col == "Hist Hit 7%":
            return "+0.0;-0.0;0.0" if col != "Hist Hit 7%" else "0"
        if "Rs" in col:
            return "#,##0;[Red]-#,##0"
        if col in ("Rank", "Qty", "Picks", "Entry Hit", "Entry Not Hit", "Target 1 Hit", "Target 2 Hit",
                   "Stop Loss Hit (before T1)", "T1 then SL (before T2)", "Still Open"):
            return "0"
        return "#,##0.00"

    def table(ws, r0, df, total_label=None):
        cols = list(df.columns)
        for j, c in enumerate(cols, 1):
            x = ws.cell(r0, j, c)
            x.font, x.fill, x.alignment, x.border = head_f, head_fill, center, bd
        ws.row_dimensions[r0].height = 36
        for i, rec in enumerate(df.itertuples(index=False), r0 + 1):
            is_total = total_label is not None and rec[0] == total_label
            for j, (c, v) in enumerate(zip(cols, rec), 1):
                if isinstance(v, float) and not np.isfinite(v):
                    v = None
                x = ws.cell(i, j, v)
                x.font, x.border = (bold_f if is_total else body_f), bd
                x.alignment = left if isinstance(v, str) else center
                if v is not None and not isinstance(v, str):
                    x.number_format = fmt(c)
                if is_total:
                    x.fill = all_fill
        return r0 + len(df)

    def title(ws, text, sub, ncol):
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncol)
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncol)
        c = ws.cell(1, 1, text)
        c.font, c.fill, c.alignment = title_f, title_fill, Alignment(vertical="center")
        ws.row_dimensions[1].height = 26
        ws.cell(2, 1, sub).font = sub_f

    wb = Workbook()
    ws = wb.active
    ws.title = "SUMMARY"
    days = list(dict.fromkeys(res.sort_values("Generated")["Report Day"]))
    sub = (f"Reports: {', '.join(days)}  |  Prices up to {asof}  |  T1 = +7%, T2 = +10% of report price  |  "
           f"% columns are out of trades entered (see RULES)")
    ncol = len(overall.columns)
    title(ws, "BACKTEST - PRIORITY vs SWING PICKS sheets (daily SwingStocks reports)", sub, ncol)
    r = 4
    ws.cell(r, 1, "1. OVERALL - all 6 reports together").font = sec_f
    r = table(ws, r + 1, overall) + 2
    for s in SHEETS:
        ws.cell(r, 1, f"{2 if s == 'PRIORITY' else 3}. {s} sheet - day by day").font = sec_f
        r = table(ws, r + 1, by_day[s], "ALL DAYS") + 2
    ws.cell(r, 1, "Colour key: green = target hit / profit, red = stop loss / loss, amber = still open, "
                  "grey = entry not hit.  Full rules on the RULES sheet.").font = sub_f
    for j, c in enumerate(overall.columns, 1):
        ws.column_dimensions[get_column_letter(j)].width = WIDTH.get(c, 11.5)
    ws.freeze_panes = "B4"
    last = ws.max_row
    for c in ("G", "I"):                           # T1 / T2 Hit %
        ws.conditional_formatting.add(f"{c}6:{c}{last}", FormulaRule(formula=[f'AND(ISNUMBER({c}6),{c}6>=0.5)'], fill=green))
    for j, c in enumerate(overall.columns, 1):
        if c.startswith(("Avg P&L", "Total P&L")):
            L = get_column_letter(j)
            ws.conditional_formatting.add(f"{L}6:{L}{last}", FormulaRule(formula=[f'AND(ISNUMBER({L}6),{L}6>0)'], fill=green))
            ws.conditional_formatting.add(f"{L}6:{L}{last}", FormulaRule(formula=[f'AND(ISNUMBER({L}6),{L}6<0)'], fill=red))

    for s in SHEETS:
        d = res[res["Sheet"] == s].sort_values(["Generated", "Rank"])[TRADE_COLS]
        w = wb.create_sheet(f"{s} TRADES")
        title(w, f"{s} sheet - every stock, trade by trade",
              f"{len(d)} stocks from {len(days)} reports  |  Prices up to {asof}  |  levels copied from the report", len(TRADE_COLS))
        end = table(w, 4, d)
        w.freeze_panes = "E5"
        w.auto_filter.ref = f"A4:{get_column_letter(len(TRADE_COLS))}{end}"
        for j, c in enumerate(TRADE_COLS, 1):
            L = get_column_letter(j)
            w.column_dimensions[L].width = WIDTH.get(c, 11)
            rng = f"{L}5:{L}{end}"
            if c in ("Entry Hit?", "T1 Hit?", "T2 Hit?"):
                w.conditional_formatting.add(rng, FormulaRule(formula=[f'{L}5="YES"'], fill=green))
                w.conditional_formatting.add(rng, FormulaRule(formula=[f'{L}5="NO"'], fill=grey))
            if c == "SL Hit?":
                w.conditional_formatting.add(rng, FormulaRule(formula=[f'{L}5="YES"'], fill=red))
            if c == "Result":
                w.conditional_formatting.add(rng, FormulaRule(formula=[f'{L}5="TARGET 2 HIT"'], fill=dgreen))
                w.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT({L}5,12)="TARGET 1 HIT"'], fill=green))
                w.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT({L}5,4)="STOP"'], fill=red))
                w.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT({L}5,4)="OPEN"'], fill=amber))
                w.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT({L}5,5)="ENTRY"'], fill=grey))
            if c.startswith(("P&L", "Max")):
                w.conditional_formatting.add(rng, FormulaRule(formula=[f'AND(ISNUMBER({L}5),{L}5>0)'], fill=green))
                w.conditional_formatting.add(rng, FormulaRule(formula=[f'AND(ISNUMBER({L}5),{L}5<0)'], fill=red))

    w = wb.create_sheet("RULES")
    title(w, "How the backtest was done", f"Prices up to {asof}", 2)
    table(w, 4, pd.DataFrame(RULES, columns=["Item", "Rule"]))
    w.column_dimensions["A"].width, w.column_dimensions["B"].width = 26, 120
    for ws_ in wb.worksheets:
        ws_.sheet_view.showGridLines = False
        ws_.page_setup.orientation = "landscape"
        ws_.sheet_properties.pageSetUpPr.fitToPage = True
    path = Path(path)
    try:
        wb.save(path)
    except PermissionError:                          # file is open in Excel
        path = path.with_name(f"{path.stem}_{datetime.now():%H%M}.xlsx")
        wb.save(path)
    return path


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("reports", nargs="*", help="report files (default: every SwingStocks_* report in ../output/daily)")
    args = ap.parse_args()
    paths = [Path(p) for p in args.reports] or sorted(REPORT_DIR.glob("SwingStocks_*.xlsx"))
    paths = [p for p in paths if not p.name.startswith("~$")]
    if not paths:
        sys.exit(f"No SwingStocks reports found in {REPORT_DIR}")
    picks = load(paths).sort_values(["Generated", "Sheet", "Rank"], kind="stable")
    print(f"{len(picks)} rows ({(picks['Sheet'] == 'PRIORITY').sum()} PRIORITY, "
          f"{(picks['Sheet'] == 'SWING PICKS').sum()} SWING PICKS) from {len(paths)} report(s)")
    syms = sorted(picks["Stock"].unique())
    hist = fetch_history(syms, picks["Generated"].min() - pd.Timedelta(days=7))
    print(f"price history for {len(hist)}/{len(syms)} stocks")
    rows = []
    for r in picks.to_dict("records"):
        d = hist.get(r["Stock"])
        if d is None:
            rows.append(evaluate(r, None, np.nan, None))
            continue
        td = trade_day(d.index, r["Generated"])
        rows.append(evaluate(r, d.loc[td:] if td is not None else None,
                             float(d["Close"].iloc[-1]), d.index[-1].date()))
    res = pd.DataFrame(rows)
    overall, by_day = summary_tables(res)
    last_bar = max(d.index[-1] for d in hist.values()).date()
    asof = f"{last_bar:%d-%b-%Y}" + (f" (intraday, fetched {datetime.now():%H:%M} IST)"
                                     if last_bar == datetime.now().date() and datetime.now().time() < time(15, 30)
                                     else " close")
    days = list(dict.fromkeys(res.sort_values("Generated")["Report Day"]))
    tag = lambda s: s.split()[1].replace("th", "").replace("st", "").replace("nd", "").replace("rd", "")
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    path = write(res, overall, by_day,
                 OUT_DIR / f"Backtest_Priority_SwingPicks_{tag(days[0])}-{tag(days[-1])}_{days[-1].split()[-1]}.xlsx", asof)
    pd.set_option("display.width", 250)
    show = ["Group", "Picks", "Entry Hit", "Target 1 Hit", "T1 Hit %", "Target 2 Hit", "T2 Hit %",
            "Stop Loss Hit (before T1)", "SL Hit %", "Still Open", "Avg P&L % (book at T1)"]
    print(overall[show].to_string(index=False))
    for s in SHEETS:
        print(f"\n{s}\n" + by_day[s][show].to_string(index=False))
    print("\nSaved:", path)


if __name__ == "__main__":
    main()
