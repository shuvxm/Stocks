"""Swing picks report for tomorrow - clean formatted workbook."""
import sys
from pathlib import Path
import pandas as pd

BASE = Path(__file__).resolve().parent
SRC = sorted((BASE / "output" / "daily").glob("NIFTY_Swing_*.xlsx"))[-1]
print("Source:", SRC.name)
x = pd.ExcelFile(SRC)
P = pd.read_excel(x, "SWING PICKS", header=3)
W = pd.read_excel(x, "WATCHLIST", header=3)

BUY_COLS = ["Rank", "Stock", "Setup", "Action", "Swing Score", "Grade",
            "Last Close", "Entry Low", "Entry High", "Stop Loss",
            "Target 1", "Target 2"]
WHY_COLS = ["Rank", "Stock", "Swing Score", "Why Selected", "Risk Flags",
            "Hist Hit 7%", "Hist Hit 10%", "Hist Samples"]
IND_COLS = ["Rank", "Stock", "Setup", "Swing Score", "Last Close", "RSI14",
            "ADX14", "ATR%", "RVOL", "1M%", "3M%", "vs NIFTY 1M%",
            "52W High Dist%", "EMA20 Dist%", "20D High Dist%",
            "Fundamental", "Next Earnings"]


def style(ws, title):
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.formatting.rule import ColorScaleRule, FormulaRule
    nc = ws.max_column
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=nc)
    c = ws.cell(1, 1, title)
    c.font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor="1F4E78")
    ws.row_dimensions[1].height = 26
    ws.freeze_panes = "A3"
    ws.auto_filter.ref = f"A2:{get_column_letter(nc)}{ws.max_row}"
    for idx in range(1, nc + 1):
        letter = get_column_letter(idx)
        vals = [str(ws.cell(r, idx).value or "") for r in range(1, min(60, ws.max_row + 1))]
        ws.column_dimensions[letter].width = min(44, max(12, max(len(v) for v in vals) + 2))
    thin = Side(style="thin", color="B4C6E7")
    bd = Border(left=thin, right=thin, top=thin, bottom=thin)
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=nc):
        for cell in row:
            cell.font = Font(name="Calibri", size=10)
            cell.border = bd
            cell.alignment = Alignment(vertical="center", wrap_text=True)
    for cell in ws[2]:
        cell.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="2E75B6")
        cell.alignment = Alignment(vertical="center", horizontal="center", wrap_text=True)
    ws.row_dimensions[2].height = 30
    for w2 in [ws]:
        w2.sheet_properties.pageSetUpPr.fitToPage = True


def main():
    out = SRC.parent / "Swing_Tomorrow_2026-10-07.xlsx"
    with pd.ExcelWriter(out, engine="openpyxl") as wr:
        P[[c for c in BUY_COLS if c in P]].to_excel(wr, sheet_name="BUY LIST", index=False, startrow=1)
        P[[c for c in WHY_COLS if c in P]].to_excel(wr, sheet_name="WHY PICKED", index=False, startrow=1)
        P[[c for c in IND_COLS if c in P]].to_excel(wr, sheet_name="PICKS INDICATORS", index=False, startrow=1)
        W[[c for c in BUY_COLS if c in W]].to_excel(wr, sheet_name="WATCHLIST", index=False, startrow=1)
        W[[c for c in IND_COLS if c in W]].to_excel(wr, sheet_name="WATCH INDICATORS", index=False, startrow=1)
        style(wr.book["BUY LIST"], "SWING BUY LIST for Wed 07-Oct-2026 | NIFTY regime: BEAR | Targets +7%/+10% | Qty = Rs 1L capital, 1% risk")
        style(wr.book["WHY PICKED"], "WHY EACH STOCK WAS PICKED + RISK FLAGS + HISTORY HIT-RATE")
        style(wr.book["PICKS INDICATORS"], "ALL INDICATORS - 15 SWING PICKS (point-in-time)")
        style(wr.book["WATCHLIST"], "WATCHLIST - 20 stocks (score 50+, buy only on trigger/confirmation)")
        style(wr.book["WATCH INDICATORS"], "ALL INDICATORS - 20 WATCHLIST")
        from openpyxl.formatting.rule import ColorScaleRule
        for sh, col in (("BUY LIST", "E"), ("PICKS INDICATORS", "D"), ("WATCHLIST", "E")):
            if sh in wr.book.sheetnames:
                w2 = wr.book[sh]
                w2.conditional_formatting.add(f"{col}3:{col}{w2.max_row}", ColorScaleRule(
                    start_type="num", start_value=50, start_color="F8696B",
                    mid_type="num", mid_value=65, mid_color="FFEB84",
                    end_type="num", end_value=85, end_color="63BE7B"))
        ws = wr.book["BUY LIST"]
        ws.column_dimensions["D"].width = 40
        ws2 = wr.book["WHY PICKED"]
        ws2.column_dimensions["D"].width = 46
        ws2.column_dimensions["E"].width = 30
    print("Saved:", out)


main()
