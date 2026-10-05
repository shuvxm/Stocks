"""NIFTY Swing Analyst v2  -  find a few high-quality swing candidates (target +7% to +10%).

Install : pip install yfinance pandas numpy openpyxl requests
Run     : python nifty_swing_analyst_v2.py
Options : --universe NIFTY50|NIFTY100|NIFTY200|NIFTY500   (default NIFTY200)
          --no-news          skip news lookup (faster)
          --keep-partial     use today's unfinished candle if market is open
          --picks 10         number of picks in the main sheet
          --min-score 60     minimum Swing Score for a pick

Output  : output/daily/NIFTY_Swing_<date_time>.xlsx
Sheets  : SWING PICKS | WATCHLIST | ALL INDICATORS | HOW TO READ

Pipeline
  1. Batch-download 2y daily data for the whole universe (fast, with fallbacks)
  2. Compute indicators (Wilder ATR/ADX/RSI), detect setups, hard-filter, score
  3. Enrich only the best ~45 with live price, news, earnings date, fundamentals
  4. Final score -> picks + watchlist -> formatted Excel

This is a screening tool, not financial advice. No screen can guarantee a +7-10% move.
"""
import argparse
import math
import time
import xml.etree.ElementTree as ET
from collections import Counter
from datetime import date, datetime
from io import StringIO
from pathlib import Path
from urllib.parse import quote_plus

import numpy as np
import pandas as pd
import requests
import yfinance as yf
from openpyxl.formatting.rule import ColorScaleRule, FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ----------------------------------------------------------------------------
# CONFIG
# ----------------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "output" / "daily"
CACHE_DIR = BASE_DIR / "output" / "cache"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
CACHE_DIR.mkdir(parents=True, exist_ok=True)

UNIVERSE = "NIFTY200"
HISTORY_PERIOD = "2y"
MIN_HISTORY = 220
CHUNK = 40                       # tickers per yfinance batch call

MIN_AVG_VOLUME = 100_000         # shares/day (20d avg)
MIN_AVG_VALUE = 20_000_000       # Rs 2 Cr traded/day (20d avg)

MIN_SCORE = 60                   # score needed for SWING PICKS
WATCH_SCORE = 50                 # score needed for WATCHLIST
MAX_PICKS = 15
MAX_WATCH = 20
SHORTLIST = 45                   # how many get news/fundamental enrichment

# Profit goal: user wants 7-10% swing moves
TARGET1_PCT = 7.0
TARGET2_PCT = 10.0
HOLD_DAYS = 15                   # ~3 weeks; used for historical hit-rate

# Risk / sizing
ACCOUNT_CAPITAL = 100_000
RISK_PER_TRADE = 0.01
MAX_POSITION_PERCENT = 0.20
STOP_ATR = 1.8
STOP_MIN_PCT = 3.0               # stop is at least this far below entry
STOP_MAX_PCT = 6.5               # ...and never further than this
MIN_RR_T2 = 1.5                  # reward:risk to Target 2

# Hard filters
RSI_MIN, RSI_MAX = 52, 78
ADX_MIN = 18
ATR_PCT_MIN, ATR_PCT_MAX = 1.3, 6.5   # stock must be able to move 7-10% in ~3 weeks
MAX_EXT_EMA20 = 10.0             # % above EMA20 = too extended to chase

NIFTY = "^NSEI"
NSE_URLS = {
    "NIFTY50": "https://nsearchives.nseindia.com/content/indices/ind_nifty50list.csv",
    "NIFTY100": "https://nsearchives.nseindia.com/content/indices/ind_nifty100list.csv",
    "NIFTY200": "https://nsearchives.nseindia.com/content/indices/ind_nifty200list.csv",
    "NIFTY500": "https://nsearchives.nseindia.com/content/indices/ind_nifty500list.csv",
    "MIDCAP150": "https://nsearchives.nseindia.com/content/indices/ind_niftymidcap150list.csv",
    "SMALLCAP250": "https://nsearchives.nseindia.com/content/indices/ind_niftysmallcap250list.csv",
    "MICROCAP250": "https://nsearchives.nseindia.com/content/indices/ind_niftymicrocap250_list.csv",
    # Composite: union of all of the above (~1250 unique tradeable stocks).
    # True 2700 needs NSE's CM master file which blocks bots; NSE_ALL covers
    # everything liquid enough to swing-trade. Àny --extra symbol is always added.
    "NSE_ALL": "COMPOSITE:NIFTY500,MIDCAP150,SMALLCAP250,MICROCAP250",
}
NSE_ANN = "https://www.nseindia.com/companies-listing/corporate-filings-announcements?symbol={}&tabIndex=equity"
FALLBACK = ("ADANIENT ADANIPORTS APOLLOHOSP ASIANPAINT AXISBANK BAJAJ-AUTO BAJFINANCE BAJAJFINSV BEL "
            "BHARTIARTL CIPLA COALINDIA DRREDDY EICHERMOT ETERNAL GRASIM HCLTECH HDFCBANK HDFCLIFE "
            "HEROMOTOCO HINDALCO HINDUNILVR ICICIBANK INDUSINDBK INFY ITC JIOFIN JSWSTEEL KOTAKBANK LT "
            "M&M MARUTI NESTLEIND NTPC ONGC POWERGRID RELIANCE SBILIFE SBIN SHRIRAMFIN SUNPHARMA "
            "TATACONSUMER TATASTEEL TCS TECHM TITAN TRENT ULTRACEMCO WIPRO").split()

POS_WORDS = ["order win", "wins order", "bags order", "secures order", "new order", "contract", "upgrade",
             "buy rating", "target raised", "raises target", "record profit", "profit jumps", "profit rises",
             "profit surges", "net profit up", "beats estimates", "strong results", "strong q", "buyback",
             "bonus", "dividend", "acquisition", "expansion", "approval", "partnership", "outperform"]
NEG_WORDS = ["downgrade", "sell rating", "fraud", "probe", "investigation", "raid", "penalty", "sebi notice",
             "sebi order", "net loss", "loss widens", "profit falls", "profit drops", "profit declines",
             "misses estimates", "weak results", "resigns", "default", "lawsuit", "ban ", "recall", "target cut",
             "cuts target", "plunge", "slump", "tumbles", "insolvency", "pledge", "auditor"]

REQ = ["Close", "Volume", "EMA20", "EMA50", "EMA200", "RSI14", "MACD", "MACDSignal", "MACDHist", "ATR14",
       "ATRpct", "ADX14", "RVOL", "OBV", "OBVEMA20", "High20P", "High50P", "High252", "Low10", "RET1",
       "RET5", "RET21", "RET63", "CLV", "DistEMA20", "StochK", "VolSMA50", "EMA50Slope"]


# ----------------------------------------------------------------------------
# HELPERS / DATA
# ----------------------------------------------------------------------------
def fnum(x, default=np.nan):
    try:
        if isinstance(x, str):
            x = x.replace(",", "").replace("%", "").strip()
        v = float(x)
        return v if np.isfinite(v) else default
    except Exception:
        return default


def sym(t):
    return str(t).replace(".NS", "").upper()


def _fetch_index_csv(url):
    """Download one NSE index csv, return (tickers, sectors, names). Empty on failure."""
    try:
        s = requests.Session()
        s.headers.update({"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
                          "Referer": "https://www.nseindia.com/", "Accept": "text/csv,*/*"})
        try:
            s.get("https://www.nseindia.com", timeout=10)
        except Exception:
            pass
        r = s.get(url, timeout=25)
        r.raise_for_status()
        df = pd.read_csv(StringIO(r.text))
        df.columns = [str(c).strip() for c in df.columns]
        col = next(c for c in df.columns if c.upper() == "SYMBOL")
        ind = next((c for c in df.columns if c.upper() == "INDUSTRY"), None)
        nm = next((c for c in df.columns if c.upper().startswith("COMPANY")), None)
        df[col] = df[col].astype(str).str.strip().str.upper()
        tickers = sorted(set(df[col].dropna() + ".NS"))
        sectors = dict(zip(df[col], df[ind])) if ind else {}
        names = dict(zip(df[col], df[nm])) if nm else {}
        return tickers, sectors, names
    except Exception as e:
        print(f"  index fetch failed ({url.split('/')[-1]}): {e}")
        return [], {}, {}


def get_universe(name):
    """Returns (tickers, {symbol: industry}, {symbol: company name}). Caches the NSE csv locally."""
    # Composite universe: union of several index lists (NSE_ALL).
    if isinstance(NSE_URLS.get(name), str) and NSE_URLS[name].startswith("COMPOSITE:"):
        parts = NSE_URLS[name].split(":", 1)[1].split(",")
        tickers, sectors, names = [], {}, {}
        seen = set()
        for p in parts:
            p = p.strip()
            cache = CACHE_DIR / f"{p}.csv"
            if cache.exists():
                try:
                    import csv as _csv
                    with open(cache, encoding="utf-8") as f:
                        rd = _csv.DictReader(f)
                        cols = {k.strip().upper(): k for k in (rd.fieldnames or [])}
                        for row in rd:
                            sy = str(row.get(cols.get("SYMBOL", "Symbol"), "")).strip().upper()
                            if sy and sy + ".NS" not in seen:
                                seen.add(sy + ".NS")
                                tickers.append(sy + ".NS")
                                if "INDUSTRY" in cols:
                                    sectors[sy] = row.get(cols["INDUSTRY"], "")
                                for ck in cols:
                                    if ck.startswith("COMPANY"):
                                        names[sy] = row.get(cols[ck], "")
                                        break
                    continue
                except Exception:
                    pass
            t, se, na = _fetch_index_csv(NSE_URLS[p])
            for x in t:
                if x not in seen:
                    seen.add(x)
                    tickers.append(x)
            sectors.update(se)
            names.update(na)
        # also merge custom list: watchlist.txt (one NSE symbol per line) if present
        wl = BASE_DIR / "watchlist.txt"
        if wl.exists():
            for line in wl.read_text(encoding="utf-8").splitlines():
                sy = line.strip().upper().removesuffix(".NS")
                if sy and sy + ".NS" not in seen:
                    seen.add(sy + ".NS")
                    tickers.append(sy + ".NS")
        tickers.sort()
        print(f"Composite universe {name}: {len(tickers)} unique stocks from {parts}"
              + (" + watchlist.txt" if wl.exists() else ""))
        return tickers, sectors, names
    cache = CACHE_DIR / f"{name}.csv"
    text = None
    try:
        s = requests.Session()
        s.headers.update({"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
                          "Referer": "https://www.nseindia.com/", "Accept": "text/csv,*/*"})
        try:
            s.get("https://www.nseindia.com", timeout=10)
        except Exception:
            pass
        r = s.get(NSE_URLS[name], timeout=25)
        r.raise_for_status()
        text = r.text
        cache.write_text(text, encoding="utf-8")
    except Exception as e:
        print("NSE universe download failed:", e)
        if cache.exists():
            print("Using cached universe file:", cache)
            text = cache.read_text(encoding="utf-8")
    if text:
        try:
            df = pd.read_csv(StringIO(text))
            df.columns = [str(c).strip() for c in df.columns]
            col = next(c for c in df.columns if c.upper() == "SYMBOL")
            ind = next((c for c in df.columns if c.upper() == "INDUSTRY"), None)
            nm = next((c for c in df.columns if c.upper().startswith("COMPANY")), None)
            df[col] = df[col].astype(str).str.strip().str.upper()
            tickers = sorted(set(df[col].dropna() + ".NS"))
            sectors = dict(zip(df[col], df[ind])) if ind else {}
            names = dict(zip(df[col], df[nm])) if nm else {}
            return tickers, sectors, names
        except Exception as e:
            print("Universe parse failed:", e)
    print("!! Using built-in NIFTY50 fallback list (results limited to these 50 stocks)")
    return [s + ".NS" for s in FALLBACK], {}, {}


COLS = ["Open", "High", "Low", "Close", "Volume"]


def clean(d):
    if d is None or len(d) == 0:
        return None
    if isinstance(d.columns, pd.MultiIndex):
        d = d.copy()
        d.columns = d.columns.get_level_values(0)
    if not all(c in d.columns for c in COLS):
        return None
    d = d[COLS].apply(pd.to_numeric, errors="coerce").dropna(subset=["Open", "High", "Low", "Close"])
    d["Volume"] = d["Volume"].fillna(0)
    d = d[~d.index.duplicated(keep="last")].sort_index()
    return d if len(d) else None


def history_single(ticker):
    for _ in range(2):
        try:
            d = clean(yf.Ticker(ticker).history(period=HISTORY_PERIOD, interval="1d", auto_adjust=True))
            if d is not None:
                return d
        except Exception:
            pass
        time.sleep(1)
    return None


def download_all(tickers):
    """Batch download (much faster and more reliable than one call per ticker)."""
    data = {}
    for i in range(0, len(tickers), CHUNK):
        chunk = tickers[i:i + CHUNK]
        raw = None
        for _ in range(3):
            try:
                raw = yf.download(chunk, period=HISTORY_PERIOD, interval="1d", auto_adjust=True,
                                  progress=False, threads=True, group_by="ticker")
                if raw is not None and not raw.empty:
                    break
            except Exception as e:
                print("\nbatch error:", e)
            time.sleep(2)
        if raw is not None and not raw.empty:
            for t in chunk:
                try:
                    if isinstance(raw.columns, pd.MultiIndex):
                        if t not in raw.columns.get_level_values(0):
                            continue
                        d = clean(raw[t].copy())
                    else:
                        d = clean(raw.copy())
                    if d is not None:
                        data[t] = d
                except Exception:
                    pass
        print(f"  downloaded {min(i + CHUNK, len(tickers))}/{len(tickers)}", end="\r")
    missing = [t for t in tickers if t not in data]
    if missing:
        print(f"\n  retrying {len(missing)} missing tickers one by one...")
        for t in missing[:60]:
            d = history_single(t)
            if d is not None:
                data[t] = d
    return data


def drop_partial(d):
    """If the NSE market is open, today's candle is incomplete (volume looks tiny). Drop it."""
    try:
        now = pd.Timestamp.now(tz="Asia/Kolkata")
        mins = now.hour * 60 + now.minute
        if now.weekday() < 5 and 9 * 60 + 15 <= mins <= 15 * 60 + 45 and pd.Timestamp(d.index[-1]).date() == now.date():
            return d.iloc[:-1]
    except Exception:
        pass
    return d


# ----------------------------------------------------------------------------
# INDICATORS
# ----------------------------------------------------------------------------
def wilder(s, n=14):
    return s.ewm(alpha=1 / n, adjust=False, min_periods=n).mean()


def indicators(d):
    x = d.copy()
    c, h, l, v = x.Close, x.High, x.Low, x.Volume
    for n in (20, 50, 200):
        x[f"EMA{n}"] = c.ewm(span=n, adjust=False).mean()
    for n in (1, 5, 10, 21, 63):
        x[f"RET{n}"] = c.pct_change(n) * 100
    x["EMA50Slope"] = (x.EMA50 / x.EMA50.shift(10) - 1) * 100
    x["DistEMA20"] = (c / x.EMA20 - 1) * 100

    delta = c.diff()
    ag, al = wilder(delta.clip(lower=0)), wilder((-delta).clip(lower=0))
    x["RSI14"] = (100 - 100 / (1 + ag / al.replace(0, np.nan))).where(al > 0, 100).where(al.notna())

    e12, e26 = c.ewm(span=12, adjust=False).mean(), c.ewm(span=26, adjust=False).mean()
    x["MACD"] = e12 - e26
    x["MACDSignal"] = x.MACD.ewm(span=9, adjust=False).mean()
    x["MACDHist"] = x.MACD - x.MACDSignal
    x["MACDHistChg"] = x.MACDHist.diff()

    pc = c.shift(1)
    tr = pd.concat([h - l, (h - pc).abs(), (l - pc).abs()], axis=1).max(axis=1)
    atr = wilder(tr)
    x["ATR14"] = atr
    x["ATRpct"] = atr / c * 100
    up, dn = h.diff(), -l.diff()
    pdm = pd.Series(np.where((up > dn) & (up > 0), up, 0.0), index=x.index)
    mdm = pd.Series(np.where((dn > up) & (dn > 0), dn, 0.0), index=x.index)
    pdi = 100 * wilder(pdm) / atr.replace(0, np.nan)
    mdi = 100 * wilder(mdm) / atr.replace(0, np.nan)
    x["PlusDI"], x["MinusDI"] = pdi, mdi
    x["ADX14"] = wilder(100 * (pdi - mdi).abs() / (pdi + mdi).replace(0, np.nan))

    mid, sd = c.rolling(20).mean(), c.rolling(20).std()
    upper, lower = mid + 2 * sd, mid - 2 * sd
    x["BBpctB"] = (c - lower) / (upper - lower).replace(0, np.nan)
    x["BBwidth"] = (upper - lower) / mid.replace(0, np.nan) * 100
    sq = (x.BBwidth <= x.BBwidth.rolling(120, min_periods=60).quantile(0.25)).astype(float)
    x["BBsqRecent"] = sq.rolling(10, min_periods=1).max()

    lo14, hi14 = l.rolling(14).min(), h.rolling(14).max()
    x["StochK"] = 100 * (c - lo14) / (hi14 - lo14).replace(0, np.nan)
    x["StochD"] = x.StochK.rolling(3).mean()

    x["VolSMA50"] = v.rolling(50).mean()
    x["RVOL"] = v / v.shift(1).rolling(20).mean().replace(0, np.nan)      # vs PRIOR 20-day average
    x["OBV"] = (np.sign(c.diff()).fillna(0) * v).cumsum()
    x["OBVEMA20"] = x.OBV.ewm(span=20, adjust=False).mean()
    upv = v.where(c > pc, 0).rolling(10).sum()
    dnv = v.where(c < pc, 0).rolling(10).sum()
    x["UDVol"] = upv / dnv.replace(0, np.nan)

    x["High20P"] = h.shift(1).rolling(20).max()                           # PRIOR 20d high (true breakout level)
    x["High50P"] = h.shift(1).rolling(50).max()
    x["High252"] = h.rolling(252, min_periods=200).max()
    x["Low10"] = l.rolling(10).min()
    x["Dist20"] = (c / x.High20P - 1) * 100
    x["Dist52W"] = (c / x.High252 - 1) * 100
    x["CLV"] = (c - l) / (h - l).replace(0, np.nan)                       # close location in day's range

    fwd = h[::-1].rolling(HOLD_DAYS, min_periods=HOLD_DAYS).max()[::-1].shift(-1)
    x["FwdMaxPct"] = (fwd / c - 1) * 100                                   # best gain in next HOLD_DAYS
    return x


def hist_stats(x):
    """How often did THIS stock gain +7% / +10% within HOLD_DAYS after similar uptrend conditions?"""
    mask = ((x.Close > x.EMA50) & (x.EMA20 > x.EMA50) & (x.EMA50 > x.EMA200)
            & x.RSI14.between(50, 78) & (x.RVOL >= 0.9))
    f = x.FwdMaxPct[mask].dropna()
    if len(f) < 10:
        return np.nan, np.nan, len(f)
    return (f >= TARGET1_PCT).mean() * 100, (f >= TARGET2_PCT).mean() * 100, len(f)


def market_regime(keep_partial=False):
    d = history_single(NIFTY)
    if d is not None and not keep_partial:
        d = drop_partial(d)
    if d is None or len(d) < MIN_HISTORY:
        return {"regime": "UNKNOWN", "price": np.nan, "r21": np.nan, "r63": np.nan}
    r = indicators(d).iloc[-1]
    s = sum([r.Close > r.EMA20, r.EMA20 > r.EMA50, r.EMA50 > r.EMA200, r.RSI14 >= 50, r.ADX14 >= 20])
    regime = "BULL" if s >= 4 else "BEAR" if s <= 1 else "NEUTRAL"
    return {"regime": regime, "price": float(r.Close), "r21": float(r.RET21), "r63": float(r.RET63)}


# ----------------------------------------------------------------------------
# SETUP / FILTER / SCORE / LEVELS
# ----------------------------------------------------------------------------
def classify(r):
    if r.Close > r.High20P and r.RVOL >= 1.3 and r.CLV >= 0.6:
        return "BREAKOUT"
    if 0.97 * r.High20P <= r.Close <= r.High20P and r.Close > r.EMA20 and r.BBsqRecent >= 1:
        return "NEAR BREAKOUT"
    if -1.5 <= r.DistEMA20 <= 2.5 and r.EMA20 > r.EMA50 and 45 <= r.RSI14 <= 65 and r.StochK <= 75:
        return "EMA20 PULLBACK"
    if r.RET5 >= 2 and r.RET21 >= 5 and r.RSI14 >= 58 and r.RVOL >= 1.2:
        return "MOMENTUM"
    if r.Close > r.EMA20 > r.EMA50 > r.EMA200 and r.MACD > r.MACDSignal and r.ADX14 >= ADX_MIN:
        return "TREND CONTINUATION"
    return "WATCH"


RVOL_MIN = {"BREAKOUT": 1.3, "NEAR BREAKOUT": 0.8, "EMA20 PULLBACK": 0.5, "MOMENTUM": 1.2, "TREND CONTINUATION": 0.9}


def passes(r, m, s):
    if not (r.Close > r.EMA50 and r.Close > r.EMA200 and r.EMA20 > r.EMA50):
        return False
    rsi_min = 45 if s == "EMA20 PULLBACK" else RSI_MIN
    if not (rsi_min <= r.RSI14 <= RSI_MAX):
        return False
    adx_min = 15 if s in ("BREAKOUT", "NEAR BREAKOUT") else ADX_MIN
    if r.ADX14 < adx_min or r.RVOL < RVOL_MIN.get(s, 1.0):
        return False
    if not (ATR_PCT_MIN <= r.ATRpct <= ATR_PCT_MAX):
        return False
    if r.DistEMA20 > MAX_EXT_EMA20 or r.RET1 > 8 or r.RET5 > 15:
        return False
    if m["regime"] == "BEAR" and np.isfinite(m["r21"]) and r.RET21 < m["r21"] + 5:
        return False
    return True


def pts(cond, n):
    return n if bool(cond) else 0


def score(r, m, h7):
    why = []
    # TREND (20)
    t = pts(r.Close > r.EMA20, 4) + pts(r.EMA20 > r.EMA50, 5) + pts(r.EMA50 > r.EMA200, 5) \
        + pts(r.Close > r.EMA200, 3) + pts(r.EMA50Slope > 0, 3)
    # MOMENTUM (20)
    mo = (6 if 55 <= r.RSI14 <= 72 else 3 if (48 <= r.RSI14 < 55 or 72 < r.RSI14 <= 78) else 0)
    mo += pts(r.MACD > r.MACDSignal, 4) + pts(r.MACDHist > 0 and r.MACDHistChg > 0, 3) \
        + pts(r.RET5 > 1, 3) + pts(r.RET21 > 4, 4)
    # VOLUME (20)
    vo = 10 if r.RVOL >= 2 else 8 if r.RVOL >= 1.5 else 5 if r.RVOL >= 1.2 else 2 if r.RVOL >= 1 else 0
    vo += pts(r.OBV > r.OBVEMA20, 4) + pts(r.Volume > r.VolSMA50, 3) + pts(np.isfinite(r.UDVol) and r.UDVol > 1.2, 3)
    # STRUCTURE / BREAKOUT (15)
    st = pts(r.Close > r.High20P, 6) + pts(r.Close > r.High50P, 4) \
        + (3 if r.Dist52W >= -5 else 1 if r.Dist52W >= -12 else 0) + pts(r.BBsqRecent >= 1, 2)
    # RELATIVE STRENGTH vs NIFTY (10)
    rs = pts(np.isfinite(m["r21"]) and r.RET21 > m["r21"] + 2, 5) + pts(np.isfinite(m["r63"]) and r.RET63 > m["r63"] + 3, 5)
    # VOLATILITY FIT (5): enough range to reach 7-10%, not wild
    vf = 5 if 1.5 <= r.ATRpct <= 4.5 else 3
    # HISTORICAL HIT-RATE of +7% in HOLD_DAYS for this stock (10)
    hh = 3 if not np.isfinite(h7) else 10 if h7 >= 35 else 7 if h7 >= 25 else 4 if h7 >= 15 else 0
    reg = {"BULL": 3, "NEUTRAL": 1, "BEAR": -5}.get(m["regime"], 0)
    total = int(max(0, min(100, t + mo + vo + st + rs + vf + hh + reg)))

    for text, ok in [("Price>EMA20>EMA50>EMA200", r.Close > r.EMA20 > r.EMA50 > r.EMA200),
                     ("MACD bullish", r.MACD > r.MACDSignal), ("ADX>=25 strong trend", r.ADX14 >= 25),
                     (f"Volume {r.RVOL:.1f}x avg", r.RVOL >= 1.3), ("OBV rising", r.OBV > r.OBVEMA20),
                     ("Closed above prior 20D high", r.Close > r.High20P),
                     ("Near 52W high", r.Dist52W >= -5), ("BB squeeze before move", r.BBsqRecent >= 1),
                     ("Beating NIFTY 1M", np.isfinite(m["r21"]) and r.RET21 > m["r21"] + 2),
                     (f"Stock hit +{TARGET1_PCT:.0f}% in {HOLD_DAYS}d {h7:.0f}% of similar setups", np.isfinite(h7) and h7 >= 25)]:
        if ok:
            why.append(text)
    return total, why


def trade_levels(r, s):
    c, p = float(r.Close), float(r.High20P)
    if s == "NEAR BREAKOUT":                      # buy only when price clears the 20D high
        lo, hi = p * 1.001, p * 1.015
        ref = (lo + hi) / 2
    elif s == "EMA20 PULLBACK":
        lo, hi, ref = min(c, float(r.EMA20)), c * 1.005, c
    else:
        lo, hi, ref = c * 0.99, c * 1.01, c
    stop = max(ref - STOP_ATR * float(r.ATR14), float(r.Low10) * 0.995)       # tighter of ATR / swing low
    stop = min(stop, ref * (1 - STOP_MIN_PCT / 100))
    stop = max(stop, ref * (1 - STOP_MAX_PCT / 100))
    risk = ref - stop
    if risk <= 0:
        return None
    t1, t2 = ref * (1 + TARGET1_PCT / 100), ref * (1 + TARGET2_PCT / 100)
    rr1, rr2 = (t1 - ref) / risk, (t2 - ref) / risk
    if rr2 < MIN_RR_T2:
        return None
    qty = max(0, min(math.floor(ACCOUNT_CAPITAL * RISK_PER_TRADE / risk), math.floor(ACCOUNT_CAPITAL * MAX_POSITION_PERCENT / ref)))
    return dict(lo=lo, hi=hi, ref=ref, stop=stop, t1=t1, t2=t2, risk=risk, rr1=rr1, rr2=rr2, qty=qty)


def entry_status(live, row):
    lo, hi, stop, s = row["Entry Low"], row["Entry High"], row["Stop Loss"], row["Setup"]
    if live <= stop:
        return "SKIP - price already below stop"
    if live > hi * 1.005:
        return "WAIT - above entry zone, don't chase"
    if live < lo * 0.995:
        return f"WAIT - trigger not hit (buy above {lo:.2f})" if s == "NEAR BREAKOUT" else "WAIT - below entry zone, needs confirmation"
    return "BUY ZONE"


# ----------------------------------------------------------------------------
# STAGE 1: evaluate one stock from price data
# ----------------------------------------------------------------------------
def evaluate(t, d, m, sectors, names):
    if len(d) < MIN_HISTORY:
        return None, "short history"
    x = indicators(d)
    r = x.iloc[-1]
    if any(pd.isna(r.get(k)) for k in REQ):
        return None, "indicator NaN"
    if x.Volume.tail(20).mean() < MIN_AVG_VOLUME or (x.Close * x.Volume).tail(20).mean() < MIN_AVG_VALUE:
        return None, "illiquid"
    s = classify(r)
    if s == "WATCH":
        return None, "no setup pattern"
    if not passes(r, m, s):
        return None, "failed trend/momentum filters"
    lv = trade_levels(r, s)
    if lv is None:
        return None, "reward:risk too low"
    h7, h10, n = hist_stats(x)
    sc, why = score(r, m, h7)

    flags = []
    if r.RSI14 > 70: flags.append("RSI>70 (hot)")
    if r.DistEMA20 > 7: flags.append("Extended >7% above EMA20")
    if r.RET1 > 5: flags.append("Big 1D spike")
    if r.ADX14 < 20: flags.append("Weak trend (ADX<20)")
    if r.ATRpct > 5: flags.append("High volatility")
    if np.isfinite(h7) and h7 < 15: flags.append(f"Low hist. +{TARGET1_PCT:.0f}% hit-rate")
    if r.RVOL < 1: flags.append("Volume below average")

    S = sym(t)
    row = {
        "Stock": S, "Company": names.get(S, ""), "Sector": sectors.get(S, ""), "Setup": s,
        "Base Score": sc, "Last Close": float(r.Close),
        "Entry Low": lv["lo"], "Entry High": lv["hi"], "Stop Loss": lv["stop"],
        "Stop %": (lv["ref"] - lv["stop"]) / lv["ref"] * 100,
        "Target 1": lv["t1"], "Target 2": lv["t2"], "RR T1": lv["rr1"], "RR T2": lv["rr2"],
        "Risk/Share": lv["risk"], "Suggested Qty*": lv["qty"], "Position Value*": lv["qty"] * lv["ref"],
        "Hist Hit 7%": h7, "Hist Hit 10%": h10, "Hist Samples": n,
        "RVOL": r.RVOL, "RSI14": r.RSI14, "ADX14": r.ADX14, "ATR%": r.ATRpct,
        "1D%": r.RET1, "5D%": r.RET5, "1M%": r.RET21, "3M%": r.RET63,
        "vs NIFTY 1M%": r.RET21 - m["r21"] if np.isfinite(m["r21"]) else np.nan,
        "52W High Dist%": r.Dist52W, "20D High Dist%": r.Dist20, "EMA20 Dist%": r.DistEMA20,
        "EMA20": r.EMA20, "EMA50": r.EMA50, "EMA200": r.EMA200,
        "MACD": r.MACD, "MACD Signal": r.MACDSignal, "MACD Hist": r.MACDHist,
        "+DI": r.PlusDI, "-DI": r.MinusDI, "ATR14": r.ATR14,
        "OBV vs EMA20": "Rising" if r.OBV > r.OBVEMA20 else "Weak",
        "Up/Down Vol 10D": r.UDVol, "BB %B": r.BBpctB, "BB Width%": r.BBwidth,
        "BB Squeeze": "Yes" if r.BBsqRecent >= 1 else "No", "Stoch K": r.StochK, "Stoch D": r.StochD,
        "NIFTY Regime": m["regime"], "NIFTY 1M%": m["r21"], "NIFTY 3M%": m["r63"],
        "_why": why, "_flags": flags, "_ticker": t,
        "NSE Announcements": NSE_ANN.format(S),
    }
    return row, None


SMALLCAP_RELAX = dict(min_history=60, max_ext=12.0, min_value=5_000_000,
                      min_volume=20_000, rvol=0.5, rsi_max=82, adx=12,
                      atr_min=1.0, atr_max=8.0, ret1=10, ret5=20)


def evaluate_relaxed(t, d, m, sectors, names, rx=SMALLCAP_RELAX):
    """Smallcap-tolerant 2nd pass. Strict evaluate() stays the default; this is
    opt-in via --smallcap. EMA200 required only if it exists (recent listings)."""
    row, _ = evaluate(t, d, m, sectors, names)
    if row:
        row["_mode"] = "strict"
        return row, None
    if len(d) < rx["min_history"]:
        return None, "short history (even relaxed)"
    x = indicators(d)
    r = x.iloc[-1]
    if pd.isna(r.get("EMA50")) or pd.isna(r.get("RSI14")) or pd.isna(r.get("ADX14")):
        return None, "indicator NaN (even relaxed)"
    if x.Volume.tail(20).mean() < rx["min_volume"] or (x.Close * x.Volume).tail(20).mean() < rx["min_value"]:
        return None, "illiquid (even relaxed)"
    s = classify(r)
    if s == "WATCH":
        return None, "no setup pattern (even relaxed)"
    ok_trend = (r.Close > r.EMA50 and r.EMA20 > r.EMA50
                and (not np.isfinite(r.get("EMA200", np.nan)) or r.Close > r.EMA200))
    if not ok_trend:
        return None, "failed trend (even relaxed)"
    if not (45 <= r.RSI14 <= rx["rsi_max"]):
        return None, f"RSI {r.RSI14:.0f} outside 45-{rx['rsi_max']} (relaxed)"
    if r.ADX14 < rx["adx"] or r.RVOL < rx["rvol"]:
        return None, f"ADX/RVOL weak (relaxed: ADX {r.ADX14:.0f}, RVOL {r.RVOL:.2f})"
    if not (rx["atr_min"] <= r.ATRpct <= rx["atr_max"]):
        return None, "ATR% outside relaxed band"
    if r.DistEMA20 > rx["max_ext"] or r.RET1 > rx["ret1"] or r.RET5 > rx["ret5"]:
        return None, "extended/spiking (even relaxed)"
    lv = trade_levels(r, s)
    if lv is None:
        return None, "reward:risk too low"
    h7, h10, hn = hist_stats(x)
    sc, why = score(r, m, h7)
    S = sym(t)
    row = {"Stock": S, "Company": names.get(S, ""), "Sector": sectors.get(S, ""), "Setup": s,
           "Base Score": sc, "Last Close": float(r.Close),
           "Entry Low": lv["lo"], "Entry High": lv["hi"], "Stop Loss": lv["stop"],
           "Stop %": (lv["ref"] - lv["stop"]) / lv["ref"] * 100,
           "Target 1": lv["t1"], "Target 2": lv["t2"], "RR T1": lv["rr1"], "RR T2": lv["rr2"],
           "Risk/Share": lv["risk"], "Suggested Qty*": lv["qty"], "Position Value*": lv["qty"] * lv["ref"],
           "Hist Hit 7%": h7, "Hist Hit 10%": h10, "Hist Samples": hn,
           "RVOL": r.RVOL, "RSI14": r.RSI14, "ADX14": r.ADX14, "ATR%": r.ATRpct,
           "1D%": r.RET1, "5D%": r.RET5, "1M%": r.RET21, "3M%": r.RET63,
           "vs NIFTY 1M%": r.RET21 - m["r21"] if np.isfinite(m["r21"]) else np.nan,
           "52W High Dist%": r.Dist52W, "20D High Dist%": r.Dist20, "EMA20 Dist%": r.DistEMA20,
           "EMA20": r.EMA20, "EMA50": r.EMA50, "EMA200": r.EMA200,
           "MACD": r.MACD, "MACD Signal": r.MACDSignal, "MACD Hist": r.MACDHist,
           "+DI": r.PlusDI, "-DI": r.MinusDI, "ATR14": r.ATR14,
           "OBV vs EMA20": "Rising" if r.OBV > r.OBVEMA20 else "Weak",
           "Up/Down Vol 10D": r.UDVol, "BB %B": r.BBpctB, "BB Width%": r.BBwidth,
           "BB Squeeze": "Yes" if r.BBsqRecent >= 1 else "No", "Stoch K": r.StochK, "Stoch D": r.StochD,
           "NIFTY Regime": m["regime"], "NIFTY 1M%": m["r21"], "NIFTY 3M%": m["r63"],
           "_why": why, "_flags": ["RELAXED smallcap"], "_ticker": t, "_mode": "relaxed",
           "NSE Announcements": NSE_ANN.format(S)}
    return row, None


# ----------------------------------------------------------------------------
# STAGE 2: enrichment (shortlist only)
# ----------------------------------------------------------------------------
def latest_price(ticker, fallback):
    try:
        fi = yf.Ticker(ticker).fast_info
        for k in ("last_price", "lastPrice", "regular_market_price"):
            try:
                q = float(fi[k])
                if q > 0 and np.isfinite(q):
                    return q, "Live quote"
            except Exception:
                pass
    except Exception:
        pass
    return float(fallback), "Last daily close"


def fundamentals(ticker):
    out = {"Market Cap Cr": np.nan, "PE": np.nan, "ROE%": np.nan, "Debt/Equity": np.nan,
           "Revenue Growth%": np.nan, "Earnings Growth%": np.nan, "Fundamental": "N/A"}
    try:
        i = yf.Ticker(ticker).info
        mc = fnum(i.get("marketCap"))
        out["Market Cap Cr"] = mc / 1e7 if np.isfinite(mc) else np.nan
        out["PE"] = fnum(i.get("trailingPE"))
        roe = fnum(i.get("returnOnEquity")); rg = fnum(i.get("revenueGrowth")); eg = fnum(i.get("earningsGrowth"))
        de = fnum(i.get("debtToEquity"))                         # Yahoo gives this in PERCENT (150 = 1.5x)
        out["ROE%"] = roe * 100 if np.isfinite(roe) else np.nan
        out["Revenue Growth%"] = rg * 100 if np.isfinite(rg) else np.nan
        out["Earnings Growth%"] = eg * 100 if np.isfinite(eg) else np.nan
        out["Debt/Equity"] = de / 100 if np.isfinite(de) else np.nan
        checks = []
        if np.isfinite(out["ROE%"]): checks.append(out["ROE%"] >= 12)
        if np.isfinite(out["Revenue Growth%"]): checks.append(out["Revenue Growth%"] > 0)
        if np.isfinite(out["Earnings Growth%"]): checks.append(out["Earnings Growth%"] > 0)
        if np.isfinite(out["Debt/Equity"]): checks.append(out["Debt/Equity"] <= 1.5)
        out["Fundamental"] = ("Supportive" if checks and sum(checks) >= math.ceil(len(checks) * 0.6)
                              else "Mixed" if checks else "N/A")
    except Exception:
        pass
    return out


def next_earnings(ticker):
    try:
        cal = yf.Ticker(ticker).calendar
        z = None
        if isinstance(cal, dict):
            z = cal.get("Earnings Date")
            if isinstance(z, (list, tuple)):
                z = z[0] if z else None
        elif isinstance(cal, pd.DataFrame) and "Earnings Date" in cal.index:
            z = cal.loc["Earnings Date"].iloc[0]
        if z is None or pd.isna(z):
            return "", None
        dt = pd.to_datetime(z).date()
        days = (dt - date.today()).days
        return dt.isoformat(), (days if days >= 0 else None)
    except Exception:
        return "", None


def fetch_news(stock, company):
    items = []
    try:
        q = quote_plus(f'"{company or stock}" shares when:7d')
        url = f"https://news.google.com/rss/search?q={q}&hl=en-IN&gl=IN&ceid=IN:en"
        r = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=15)
        r.raise_for_status()
        for it in ET.fromstring(r.content).iter("item"):
            title = (it.findtext("title") or "").strip()
            if title:
                items.append(((it.findtext("pubDate") or "")[5:16], title))
            if len(items) >= 5:
                break
    except Exception:
        pass
    if not items:
        try:
            for n in (yf.Ticker(stock + ".NS").news or [])[:5]:
                c = n.get("content") or n
                title = c.get("title") or n.get("title")
                if title:
                    items.append((str(c.get("pubDate") or "")[:10], str(title)))
        except Exception:
            pass
    return items


def news_sentiment(items):
    """Very simple headline keyword score. Treat as a prompt to read the news, not as a verdict."""
    s, neg = 0, 0
    for _, t in items:
        tl = t.lower()
        if any(w in tl for w in NEG_WORDS):
            s -= 2; neg += 1
        elif any(w in tl for w in POS_WORDS):
            s += 1
    return max(-6, min(4, s)), neg


def enrich(row, with_news):
    t, S = row["_ticker"], row["Stock"]
    p, src = latest_price(t, row["Last Close"])
    row["Live Price"], row["Price Source"] = p, src
    row["Action"] = entry_status(p, row)
    row.update(fundamentals(t))
    ed, dte = next_earnings(t)
    row["Next Earnings"], row["Days to Earnings"] = ed, dte
    adj, avoid = 0, False
    if row["Fundamental"] == "Supportive":
        adj += 3
    if dte is not None and dte <= 5:
        adj -= 5
        row["_flags"].append(f"Earnings in {dte} day(s)")
        if dte <= 3:
            avoid = True
            row["Action"] = f"REVIEW - earnings in {dte} day(s)"
    items = fetch_news(S, row["Company"]) if with_news else []
    sent, neg = news_sentiment(items)
    row["News Sentiment"] = ("Positive" if sent >= 2 else "Negative" if sent <= -2 else "Neutral") if items else ""
    row["Recent News"] = " | ".join(f"[{d}] {tt}" for d, tt in items[:4])[:900]
    adj += sent
    if neg >= 2:
        avoid = True
        row["Action"] = "AVOID - negative news, read headlines"
        row["_flags"].append("Multiple negative headlines")
    elif neg == 1:
        row["_flags"].append("Negative headline - check news")
    row["Swing Score"] = int(max(0, min(100, row["Base Score"] + adj)))
    row["Grade"] = "A" if row["Swing Score"] >= 75 else "B" if row["Swing Score"] >= 65 else "C"
    row["_avoid"] = avoid
    row["Why Selected"] = " | ".join(row["_why"][:8])
    row["Risk Flags"] = " | ".join(row["_flags"])
    return row


# ----------------------------------------------------------------------------
# EXCEL
# ----------------------------------------------------------------------------
MAIN_COLS = ["Rank", "Stock", "Company", "Sector", "Setup", "Action", "Swing Score", "Grade", "Last Close",
             "Live Price", "Entry Low", "Entry High", "Stop Loss", "Stop %", "Target 1", "Target 2", "RR T1",
             "RR T2", "Suggested Qty*", "Position Value*", "Hist Hit 7%", "Hist Hit 10%", "Hist Samples",
             "RVOL", "RSI14", "ADX14", "ATR%", "1M%", "3M%", "vs NIFTY 1M%", "52W High Dist%", "News Sentiment",
             "Next Earnings", "Fundamental", "Why Selected", "Risk Flags", "Recent News", "NSE Announcements"]
IND_COLS = ["Rank", "Stock", "Setup", "Swing Score", "Last Close", "EMA20", "EMA50", "EMA200", "RSI14", "MACD",
            "MACD Signal", "MACD Hist", "ADX14", "+DI", "-DI", "ATR14", "ATR%", "RVOL", "OBV vs EMA20",
            "Up/Down Vol 10D", "BB %B", "BB Width%", "BB Squeeze", "Stoch K", "Stoch D", "1D%", "5D%", "1M%",
            "3M%", "20D High Dist%", "52W High Dist%", "EMA20 Dist%", "NIFTY Regime", "NIFTY 1M%", "NIFTY 3M%",
            "Market Cap Cr", "PE", "ROE%", "Debt/Equity", "Revenue Growth%", "Earnings Growth%"]
PRICE_FMT = {"Last Close", "Live Price", "Entry Low", "Entry High", "Stop Loss", "Target 1", "Target 2", "EMA20",
             "EMA50", "EMA200", "ATR14", "MACD", "MACD Signal", "MACD Hist", "Risk/Share"}
INT_FMT = {"Rank", "Swing Score", "Suggested Qty*", "Hist Samples", "Position Value*", "Market Cap Cr"}
WIDTH = {"Rank": 6, "Stock": 13, "Company": 28, "Sector": 20, "Setup": 21, "Action": 40, "Swing Score": 9,
         "Grade": 7, "Why Selected": 55, "Risk Flags": 40, "Recent News": 90, "NSE Announcements": 14,
         "Next Earnings": 13, "News Sentiment": 11, "Fundamental": 12}
WRAP = {"Why Selected", "Risk Flags", "Recent News", "Action", "Company"}
FONT = "Arial"


def style_sheet(ws, df, title, subtitle):
    ws["A1"], ws["A2"] = title, subtitle
    ws["A1"].font, ws["A2"].font = Font(name=FONT, size=16, bold=True), Font(name=FONT, italic=True, size=10)
    fill = PatternFill("solid", fgColor="1F4E78")
    for c in ws[4]:
        c.fill, c.font = fill, Font(name=FONT, color="FFFFFF", bold=True, size=10)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[4].height = 34
    ws.freeze_panes = "C5"
    last = ws.max_row
    if len(df):
        ws.auto_filter.ref = f"A4:{get_column_letter(ws.max_column)}{last}"
    for j, col in enumerate(df.columns, 1):
        ws.column_dimensions[get_column_letter(j)].width = WIDTH.get(col, 12)
        fmt = "0.00" if col in PRICE_FMT else "#,##0" if col in INT_FMT else "0.0"
        for i in range(5, last + 1):
            cell = ws.cell(i, j)
            cell.font = Font(name=FONT, size=10)
            cell.alignment = Alignment(wrap_text=col in WRAP, vertical="top")
            if isinstance(cell.value, (int, float)):
                cell.number_format = fmt
            if col == "NSE Announcements" and cell.value:
                cell.hyperlink = cell.value
                cell.value = "Open filings"
                cell.font = Font(name=FONT, size=10, color="0563C1", underline="single")
    if len(df):
        def cf(rng, formula, color):
            ws.conditional_formatting.add(rng, FormulaRule(formula=[formula], fill=PatternFill(
                start_color=color, end_color=color, fill_type="solid")))
        if "Action" in df.columns:
            a = get_column_letter(list(df.columns).index("Action") + 1)
            rng = f"{a}5:{a}{last}"
            cf(rng, f'LEFT({a}5,3)="BUY"', "C6EFCE")
            cf(rng, f'LEFT({a}5,4)="WAIT"', "FFEB9C")
            cf(rng, f'OR(LEFT({a}5,5)="AVOID",LEFT({a}5,6)="REVIEW",LEFT({a}5,4)="SKIP")', "FFC7CE")
        if "Swing Score" in df.columns:
            s = get_column_letter(list(df.columns).index("Swing Score") + 1)
            ws.conditional_formatting.add(f"{s}5:{s}{last}", ColorScaleRule(
                start_type="num", start_value=50, start_color="F8CBAD",
                mid_type="num", mid_value=65, mid_color="FFEB84",
                end_type="num", end_value=85, end_color="63BE7B"))


def write_excel(picks, watch, universe, m, funnel, path):
    stamp = f"Universe: {universe} | NIFTY regime: {m['regime']} | Generated: {datetime.now():%d-%b-%Y %H:%M} | " \
            f"Targets: +{TARGET1_PCT:.0f}% / +{TARGET2_PCT:.0f}% | *Qty sized for Rs {ACCOUNT_CAPITAL:,} capital, {RISK_PER_TRADE*100:.0f}% risk"
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        for name, rows, title in (("SWING PICKS", picks, "NIFTY SWING PICKS"),
                                  ("WATCHLIST", watch, "WATCHLIST - near-miss / needs trigger or review")):
            df = pd.DataFrame(rows).reindex(columns=MAIN_COLS)
            df.to_excel(w, sheet_name=name, index=False, startrow=3)
            ws = w.sheets[name]
            style_sheet(ws, df, title, stamp)
            if not rows:
                ws["A5"] = "No stocks qualified today. See HOW TO READ sheet for the filter funnel and how to relax filters."
        allr = picks + watch
        df = pd.DataFrame(allr).reindex(columns=IND_COLS)
        df.to_excel(w, sheet_name="ALL INDICATORS", index=False, startrow=3)
        style_sheet(w.sheets["ALL INDICATORS"], df, "ALL INDICATORS (picks + watchlist)", stamp)

        ws = w.book.create_sheet("HOW TO READ")
        lines = [
            ("HOW TO READ THIS REPORT", True),
            ("Action: BUY ZONE = live price is inside entry range now. WAIT = good setup but price is outside the zone (don't chase). "
             "REVIEW/AVOID = earnings within 3 days or several negative headlines. SKIP = already below stop.", False),
            ("Setups: BREAKOUT (close above prior 20D high on >=1.3x volume) | NEAR BREAKOUT (squeezed, within 3% of the 20D high - buy only above trigger) | "
             "EMA20 PULLBACK (uptrend dipping to EMA20) | MOMENTUM | TREND CONTINUATION.", False),
            (f"Targets are fixed at +{TARGET1_PCT:.0f}% and +{TARGET2_PCT:.0f}% from entry. Stop = tighter of {STOP_ATR} x ATR or 10-day swing low, "
             f"limited to {STOP_MIN_PCT:.0f}-{STOP_MAX_PCT:.1f}% below entry. Suggested approach: book half at Target 1, trail the rest.", False),
            (f"Hist Hit 7%/10% = % of past days (this stock, last ~2y, similar uptrend conditions) where price reached the target within {HOLD_DAYS} trading days. "
             "It ignores whether the stop was hit first, so treat it as a rough 'can this stock move that far' guide.", False),
            ("Swing Score (0-100): trend 20, momentum 20, volume 20, breakout structure 15, relative strength 10, volatility fit 5, historical hit-rate 10, "
             "market regime +3/-5, then news / fundamentals / earnings adjustments. Grade A>=75, B>=65, C otherwise.", False),
            ("News Sentiment is a simple headline keyword check. Always read the headlines (Recent News column) before trading.", False),
            ("Signals use the last COMPLETED daily candle. Live Price is only used to decide if you are still inside the entry zone.", False),
            ("Not financial advice. Position sizes are illustrative; use your own capital and risk rules. Past patterns do not guarantee future moves.", False),
            ("", False), ("FILTER FUNNEL (this run)", True),
        ] + [(f"{k}: {v}", False) for k, v in funnel.items()] + [
            ("", False),
            ("If too few stocks qualify, relax in CONFIG: MIN_SCORE (60->55), RSI_MIN, ADX_MIN, ATR_PCT_MIN, or use --universe NIFTY500.", False)]
        ws.column_dimensions["A"].width = 150
        for i, (txt, bold) in enumerate(lines, 1):
            c = ws.cell(i, 1, txt)
            c.font = Font(name=FONT, bold=bold, size=12 if bold else 10)
            c.alignment = Alignment(wrap_text=True, vertical="top")
    return path


# ----------------------------------------------------------------------------
# MAIN
# ----------------------------------------------------------------------------
def main():
    global MIN_SCORE, MAX_PICKS
    ap = argparse.ArgumentParser()
    ap.add_argument("--universe", choices=list(NSE_URLS), default=UNIVERSE)
    ap.add_argument("--no-news", action="store_true")
    ap.add_argument("--keep-partial", action="store_true")
    ap.add_argument("--picks", type=int, default=MAX_PICKS)
    ap.add_argument("--min-score", type=int, default=MIN_SCORE)
    ap.add_argument("--smallcap", action="store_true",
                    help="2nd-pass smallcap tolerance (60 bars, RVOL>=0.5, ext<=12%%). Needed for micro/smallcap alerts.")
    ap.add_argument("--watchlist", default=None,
                    help="comma list of extra NSE symbols to always scan (e.g. TFCILTD,RML,GKSL,ROLEXRINGS)")
    args = ap.parse_args()
    MIN_SCORE, MAX_PICKS = args.min_score, args.picks

    print("NIFTY SWING ANALYST v2 | Universe:", args.universe, "| smallcap mode:", args.smallcap)
    m = market_regime(args.keep_partial)
    print("NIFTY regime:", m["regime"], "| price:", m["price"])
    tickers, sectors, names = get_universe(args.universe)
    if args.watchlist:
        for w in args.watchlist.split(","):
            sy = w.strip().upper().removesuffix(".NS")
            if sy and sy + ".NS" not in set(tickers):
                tickers.append(sy + ".NS")
        print(f"+ watchlist extras -> {len(tickers)} total")
    print("Stocks in universe:", len(tickers))
    data = download_all(tickers)
    print(f"\nPrice data received for {len(data)}/{len(tickers)} stocks")

    funnel = Counter()
    funnel["1. Universe size"] = len(tickers)
    funnel["2. Price data downloaded"] = len(data)
    cands = []
    for t, d in data.items():
        if not args.keep_partial:
            d = drop_partial(d)
        fn = evaluate_relaxed if args.smallcap else evaluate
        try:
            row, why = fn(t, d, m, sectors, names)
        except Exception as e:
            row, why = None, f"error: {e}"
        if row:
            cands.append(row)
        else:
            funnel[f"   rejected - {why}"] += 1
    funnel["3. Passed technical filters"] = len(cands)
    cands.sort(key=lambda r: r["Base Score"], reverse=True)
    short = [r for r in cands if r["Base Score"] >= WATCH_SCORE - 8][:SHORTLIST]
    funnel["4. Shortlisted for news/fundamentals"] = len(short)
    print(f"Passed filters: {len(cands)} | enriching top {len(short)} (live price, news, earnings, fundamentals)...")

    final = []
    for i, r in enumerate(short, 1):
        print(f"  [{i}/{len(short)}] {r['Stock']}", end="\r")
        try:
            final.append(enrich(r, not args.no_news))
        except Exception as e:
            print("\n  enrich failed for", r["Stock"], e)
        time.sleep(0.2)
    final.sort(key=lambda r: r["Swing Score"], reverse=True)
    picks = [r for r in final if r["Swing Score"] >= MIN_SCORE and not r["_avoid"]][:MAX_PICKS]
    ids = {id(r) for r in picks}
    watch = [r for r in final if id(r) not in ids and r["Swing Score"] >= WATCH_SCORE][:MAX_WATCH]
    for i, r in enumerate(picks, 1): r["Rank"] = i
    for i, r in enumerate(watch, 1): r["Rank"] = i
    funnel["5. Final SWING PICKS"] = len(picks)
    funnel["6. WATCHLIST"] = len(watch)

    path = OUTPUT_DIR / f"NIFTY_Swing_{datetime.now():%Y-%m-%d_%H%M}.xlsx"
    write_excel(picks, watch, args.universe, m, funnel, path)
    print("\n\nFunnel:")
    for k, v in funnel.items():
        print(f"  {k}: {v}")
    print("\nExcel saved:", path)
    for r in picks[:10]:
        print(f"{r['Stock']:<12}{r['Setup']:<20}score {r['Swing Score']:<4}{r['Action']:<42}"
              f"entry {r['Entry Low']:.1f}-{r['Entry High']:.1f}  SL {r['Stop Loss']:.1f}  T1 {r['Target 1']:.1f}  T2 {r['Target 2']:.1f}")


if __name__ == "__main__":
    main()