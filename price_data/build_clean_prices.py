"""Clean daily adjusted-close file for every ticker in CM Tickers.xlsx (+ EXTRA_TICKERS).

    python build_clean_prices.py                    # 1 Dec 2022 to the last completed close
    python build_clean_prices.py --start 2024-01-01

Pulls close + adjClose from FMP, repairs known data errors (listed in KNOWN_FIXES and the
automatic one-day blip check), and writes Outputs/Clean Adjusted Close <start> to <end>.xlsx with:
    Prices   adjusted close, one column per FMP ticker, one row per trading date (filled only
             inside each ticker's own trading life: blank before listing / after delisting)
    Close    raw (unadjusted) close, same layout; truncations/drops applied, dividend fixes not
    Tickers  name, root ticker, FMP ticker, sectors, currency, first/last date, notes
    Fixes    every change made to FMP's adjClose, with the reason
plus a matching .csv of Prices. Needs FMP_API_KEY in the environment (never printed or saved).
"""
from __future__ import annotations

import argparse
import json
import os
import sys
import time
import urllib.parse
import urllib.request
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timezone
from pathlib import Path

import pandas as pd

BASE = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE))
from refresh_prices import EXTRA_TICKERS, expected_day  # noqa: E402

# Dividends (in the same units as FMP's split-adjusted close) used to rebuild adjClose
# where FMP's own dividend adjustment is wrong.
REBUILD_DIVIDENDS = {
    # Hammerson: FMP's adjusted dividends before the 30 Sep 2024 1-for-10 consolidation are
    # 10x too big (72 / 78 / 75.6p instead of 7.2 / 7.8 / 7.56p), giving three fake ~35% jumps.
    "HMSO.L": {"2023-08-24": 7.2, "2024-04-04": 7.8, "2024-08-22": 7.56, "2025-04-24": 8.07,
               "2025-09-04": 7.94, "2026-03-26": 8.56, "2026-09-03": 9.67},
    # Tokyo Electron: dividends from Sep 2023 to Mar 2026 were not adjusted for the 29 Sep 2026
    # 5:1 split (close is), so each ex-date added back ~5x the real dividend.
    "8035.T": {"2023-03-30": 56.932, "2023-09-28": 29.6, "2024-03-28": 49.0, "2024-09-27": 53.0,
               "2025-03-28": 65.4, "2025-09-29": 52.8, "2026-03-30": 72.8, "2026-09-29": 76.8},
}
# Keep only prices up to and including this date (data after it is not this security).
TRUNCATE_AFTER = {
    "AVB": ("2026-08-14", "AvalonBay merged into Vivmark (VMRK) in Aug 2026; FMP joined the merged "
                          "entity's prices onto AVB from 17 Aug, a fake -64% day. Use VMRK after this."),
}
# Drop entirely: FMP now returns a different security for this ticker.
DROP = {
    "SATS": "FMP's SATS is now Strive Series A preferred stock (listed Oct 2026), not EchoStar. "
            "No EchoStar prices available under this ticker - needs a new FMP ticker.",
}
NOTES = {
    "VMRK": "Vivmark Residential: history before Aug 2026 is Equity Residential's (EQR renamed) - "
            "confirmed as the intended predecessor.",
    "EQR": "Same prices as VMRK (EQR renamed Vivmark Residential); ends 4 Sep 2026.",
    "LUMO.HE": "Kojamo renamed Lumo; history identical to KOJAMO.HE.",
    "KOJAMO.HE": "Renamed Lumo (LUMO.HE) Mar 2026.",
    "ALTRA.ST": "History identical to NYF.ST (Nyfosa renamed Altra).",
    "NYF.ST": "Renamed Altra (ALTRA.ST) May 2026.",
    "000660.KS": "+30% on 31 Jul 2026 after -27% over 28-30 Jul; same in raw close, not changed.",
    "005930.KS": "+27% on 31 Jul 2026 after -18% over 28-30 Jul; same in raw close, not changed.",
}


def fmp(path: str, key: str, **q):
    q["apikey"] = key
    url = f"https://financialmodelingprep.com/api/v3/{path}?{urllib.parse.urlencode(q)}"
    for attempt in range(6):
        if attempt:
            time.sleep(min(20, 2 ** attempt))
        try:
            with urllib.request.urlopen(url, timeout=30) as r:
                return json.load(r)
        except Exception:  # noqa: BLE001  (never print the URL: it contains the key)
            pass
    return None


def fetch(t: str, key: str, start: str, end: str) -> pd.DataFrame | None:
    p = fmp(f"historical-price-full/{urllib.parse.quote(t, safe='')}", key, **{"from": start, "to": end})
    h = p.get("historical") if isinstance(p, dict) else None
    if not h:
        return None
    d = pd.DataFrame(h)[["date", "close", "adjClose"]]
    d["date"] = pd.to_datetime(d["date"])
    return d.dropna().drop_duplicates("date").set_index("date").sort_index()


def rebuild(close: pd.Series, divs: dict[str, float]) -> pd.Series:
    """Total-return adjusted close from raw close and dividends, anchored to the last close."""
    close = close.dropna()
    factor, out = 1.0, {}
    ex = {pd.Timestamp(k): v for k, v in divs.items()}
    dates = list(close.index)
    for i in range(len(dates) - 1, -1, -1):
        d = dates[i]
        out[d] = close[d] * factor
        if d in ex and i > 0:
            factor *= 1 - ex[d] / close[dates[i - 1]]
    return pd.Series(out).sort_index()


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--start", default="2022-12-01")
    ap.add_argument("--end", default=None)
    ap.add_argument("--out", type=Path, default=BASE / "Outputs")
    a = ap.parse_args(argv)
    key = os.environ.get("FMP_API_KEY")
    if not key:
        print("FMP_API_KEY is not set in the environment.")
        return 2
    end = a.end or expected_day(datetime.now(timezone.utc)).isoformat()
    lead = (pd.Timestamp(a.start) - pd.Timedelta(days=40)).date().isoformat()  # for ex-date lookback

    cm = pd.read_excel(BASE / "CM Tickers.xlsx").dropna(subset=["FMP Ticker"])
    cm["FMP Ticker"] = cm["FMP Ticker"].astype(str).str.strip()
    tickers = list(dict.fromkeys(list(cm["FMP Ticker"]) + list(EXTRA_TICKERS)))
    print(f"Pulling {len(tickers)} tickers, {a.start} to {end} ...")
    with ThreadPoolExecutor(6) as pool:
        data = dict(zip(tickers, pool.map(lambda t: fetch(t, key, lead, end), tickers)))
    for t in [t for t, d in data.items() if d is None]:  # slow sequential retry for failures
        time.sleep(1)
        data[t] = fetch(t, key, lead, end)

    fixes, series, closes = [], {}, {}
    for t, d in data.items():
        if d is None:
            fixes.append((t, None, None, "no data returned by FMP", None, None))
            continue
        if t in DROP:
            fixes.append((t, None, None, "dropped: " + DROP[t], None, None))
            continue
        adj = d["adjClose"].copy()
        if t in REBUILD_DIVIDENDS:
            new = rebuild(d["close"], REBUILD_DIVIDENDS[t])
            chg = (new / adj - 1).abs() > 0.001
            fixes.append((t, adj.index[chg].min().date() if chg.any() else None,
                          adj.index[chg].max().date() if chg.any() else None,
                          "adjClose rebuilt from close with corrected dividends", None, None))
            adj = new
        else:
            # one-day blips: adj/close ratio jumps for a single day then returns
            ratio = adj / d["close"]
            prev, nxt = ratio.shift(1), ratio.shift(-1)
            blip = ((ratio / prev - 1).abs() > 0.005) & ((nxt / prev - 1).abs() < 0.001)
            for day in ratio.index[blip]:
                fixed = d["close"][day] * prev[day]
                fixes.append((t, day.date(), day.date(), "one-day adjClose glitch; set to close x prior-day ratio",
                              round(adj[day], 4), round(fixed, 4)))
                adj[day] = fixed
        if t in TRUNCATE_AFTER:
            cut, why = TRUNCATE_AFTER[t]
            dropped = adj[adj.index > cut]
            if len(dropped):
                fixes.append((t, dropped.index.min().date(), dropped.index.max().date(), "truncated: " + why, None, None))
            adj = adj[adj.index <= cut]
        series[t] = adj
        cl = d["close"].copy()
        if t in TRUNCATE_AFTER:
            cl = cl[cl.index <= TRUNCATE_AFTER[t][0]]
        closes[t] = cl

    wide = pd.DataFrame(series).sort_index()
    wide = wide.loc[a.start:end]
    first = {t: s.loc[a.start:end].first_valid_index() for t, s in series.items()}
    last = {t: s.loc[a.start:end].last_valid_index() for t, s in series.items()}
    # forward-fill holidays only inside each ticker's own trading life
    for t in wide.columns:
        f, l = first[t], last[t]
        if f is not None:
            wide.loc[f:l, t] = wide.loc[f:l, t].ffill()
    wide = wide[[t for t in tickers if t in wide.columns]]
    wide.index.name = "date"
    close_wide = pd.DataFrame(closes).sort_index().reindex(wide.index)
    for t in close_wide.columns:
        f, l = first[t], last[t]
        if f is not None:
            close_wide.loc[f:l, t] = close_wide.loc[f:l, t].ffill()
    close_wide = close_wide[[t for t in wide.columns if t in close_wide.columns]]

    # ticker map: one row per FMP ticker, sectors joined if listed more than once
    info = cm.groupby("FMP Ticker", sort=False).agg(
        Name=("Name", "first"), Ticker=("Ticker", "first"), Currency=("Currency", "first"),
        Sector=("Sector", lambda s: " | ".join(dict.fromkeys(s.dropna().astype(str)))),
        **{"CM Sector": ("CM Sector", lambda s: " | ".join(dict.fromkeys(s.dropna().astype(str)))),
           "CM New Sector": ("CM New Sector", lambda s: " | ".join(dict.fromkeys(s.dropna().astype(str))))}
    ).reset_index()
    extra = [t for t in EXTRA_TICKERS if t not in set(info["FMP Ticker"])]
    rows = []
    for t in extra:
        p = fmp(f"profile/{t}", key)
        p = p[0] if isinstance(p, list) and p else {}
        rows.append({"FMP Ticker": t, "Name": p.get("companyName"), "Ticker": t.split(".")[0],
                     "Currency": p.get("currency"), "Sector": "(extra - not in CM Tickers.xlsx)"})
    info = pd.concat([info, pd.DataFrame(rows)], ignore_index=True)
    info = info[["Name", "Ticker", "FMP Ticker", "Sector", "CM Sector", "CM New Sector", "Currency"]]
    info["First date"] = info["FMP Ticker"].map(lambda t: first.get(t).date() if first.get(t) is not None else None)
    info["Last date"] = info["FMP Ticker"].map(lambda t: last.get(t).date() if last.get(t) is not None else None)
    fixed_t = {f[0] for f in fixes}
    info["Status"] = info["FMP Ticker"].map(lambda t: "dropped" if t in DROP else "no data" if data.get(t) is None
                                            else "fixed" if t in fixed_t else "ok")
    info["Notes"] = info["FMP Ticker"].map(lambda t: DROP.get(t) or (TRUNCATE_AFTER.get(t) or (None, None))[1] or NOTES.get(t))
    fx = pd.DataFrame(fixes, columns=["FMP Ticker", "From", "To", "Fix", "FMP adjClose", "New adjClose"])

    a.out.mkdir(parents=True, exist_ok=True)
    stem = f"Clean Adjusted Close {a.start} to {end}"
    xlsx = a.out / f"{stem}.xlsx"
    with pd.ExcelWriter(xlsx, engine="openpyxl") as xw:
        wide.reset_index().assign(date=lambda x: x["date"].dt.date).to_excel(xw, sheet_name="Prices", index=False)
        close_wide.reset_index().assign(date=lambda x: x["date"].dt.date).to_excel(xw, sheet_name="Close", index=False)
        info.to_excel(xw, sheet_name="Tickers", index=False)
        fx.to_excel(xw, sheet_name="Fixes", index=False)
        for ws in xw.book.worksheets:
            ws.freeze_panes = "B2"
    wide.to_csv(a.out / f"{stem}.csv")
    close_wide.to_csv(a.out / f"Close {a.start} to {end}.csv")
    print(f"Saved {xlsx}\n  {wide.shape[1]} tickers x {len(wide)} dates ({wide.index.min().date()} to {wide.index.max().date()})")
    print(f"  {len(fx)} fixes applied:")
    print(fx.to_string(index=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
