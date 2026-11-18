import json, math, os, sys, re
import numpy as np, pandas as pd, openpyxl
sys.path.insert(0, os.path.dirname(__file__))
from house import *
from matplotlib.ticker import FuncFormatter, NullLocator
from matplotlib.patches import Rectangle
from load import F as VA_FILE, sheet

HERE = os.path.dirname(__file__)
OUT = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, "house_out")
PRICE_FILE = "/root/.claude/uploads/d82c9cc5-7263-5271-8116-79a86939a805/../../../tmp/claude-0/-home-user-bee2041-p1/d82c9cc5-7263-5271-8116-79a86939a805/scratchpad/Adjusted_Close_2026-01-01_to_2026-11-16.xlsx"
PRICE_FILE = "/tmp/claude-0/-home-user-bee2041-p1/d82c9cc5-7263-5271-8116-79a86939a805/scratchpad/Adjusted_Close_2026-01-01_to_2026-11-16.xlsx"

SECTORS = {
    "Cybersecurity": ["CRWD", "PANW", "NET"],
    "Tech": ["KLAC", "LRCX", "ASML", "AMAT", "000660", "MU", "2330", "AVGO", "INTC", "ARM", "CSCO", "PLTR", "SHOP", "SPCX"],
    "Other": ["COST", "WMT", "TSLA", "MNST", "WELL", "BE"],
}
SECTOR_OF = {t: s for s, ts in SECTORS.items() for t in ts}
SHORT = {"KLAC": "KLA", "LRCX": "Lam Research", "ASML": "ASML", "AMAT": "Applied Materials", "000660": "SK Hynix",
         "MU": "Micron", "2330": "TSMC", "AVGO": "Broadcom", "INTC": "Intel", "ARM": "Arm", "CSCO": "Cisco",
         "PLTR": "Palantir", "SHOP": "Shopify", "SPCX": "SpaceX", "CRWD": "CrowdStrike", "PANW": "Palo Alto Networks",
         "NET": "Cloudflare", "COST": "Costco", "WMT": "Walmart", "TSLA": "Tesla", "MNST": "Monster Beverage",
         "WELL": "Welltower", "BE": "Bloom Energy"}
FYE = {"KLAC": 6, "LRCX": 6, "ASML": 12, "CSCO": 7, "AMAT": 10, "000660": 12, "MU": 8, "AVGO": 10, "2330": 12,
       "PLTR": 12, "COST": 8, "WMT": 1, "SPCX": 12, "TSLA": 12, "SHOP": 12, "ARM": 3, "CRWD": 1, "PANW": 7,
       "INTC": 12, "MNST": 12, "WELL": 12, "BE": 12, "NET": 12}
MONTH = ["", "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
CCY = {"000660": "KRW", "2330": "TWD", "ASML": "EUR"}
SYM = {"USD": "$", "EUR": "€", "TWD": "NT$", "KRW": ""}
REV_SCALE = {"000660": 1e9, "2330": 1e6}  # VA revenue units -> currency units; default 1e6 (m)
FMP_SYM = {"000660": "000660.KS", "2330": "2330.TW"}
YEARS = [2025, 2026, 2027, 2028, 2029]

os.makedirs(OUT, exist_ok=True)

# ------------------------------------------------------------------ data
wb = openpyxl.load_workbook(VA_FILE, data_only=True, read_only=True)
ctl = {r[0]: r[1] for r in wb["Control"].iter_rows(values_only=True) if r[0]}
SNAP = pd.Timestamp("1899-12-30") + pd.Timedelta(days=int(ctl["Snapshot date"]))
act = sheet(wb, "Actuals"); con = sheet(wb, "Consensus")
act = act[act.Period_Type == "Annual"].copy(); con = con[con.Period_Type == "Annual"].copy()
num = lambda s: pd.to_numeric(s, errors="coerce")
act["val"] = num(act.Actual_Value); act["reported"] = num(act.Value_Date).notna()
con["val"] = num(con.Consensus_Value)
con["asof"] = pd.Timestamp("1899-12-30") + pd.to_timedelta(num(con.As_Of_Date), unit="D")

def metric_of(li, t):
    l = li.lower()
    if t == "WELL" and l.startswith("normalized ffo"): return "op"
    if "eps" not in l: return "rev"
    return "op" if "operating" in l else "gaap"
for d in (act, con):
    d["metric"] = [metric_of(li, t) for li, t in zip(d.Line_Item, d.Ticker)]
c0 = con[con.Lookback_Days == 0][["Ticker", "metric", "Fiscal_Year", "val"]].rename(columns={"val": "cons0"})
A = act.merge(c0, on=["Ticker", "metric", "Fiscal_Year"])
A["value"] = A.val.where(A.reported, A.cons0)
A["status"] = np.where(A.reported, "A", "E")
A.loc[(A.Ticker == "SPCX") & (A.Fiscal_Year == 2025), "status"] = "A"   # VA actuals, no report date (pre-IPO)
V = {(r.Ticker, r.metric, r.Fiscal_Year): (r.value, r.status) for r in A.itertuples()}

def val(t, m, y): return V.get((t, m, y), (np.nan, None))[0]
def stat(t, m, y): return V.get((t, m, y), (np.nan, None))[1]
def has(t, m): return any(not np.isnan(val(t, m, y)) for y in YEARS)

def fwd_year(t):
    """First fiscal year ending more than three months after the snapshot."""
    for y in YEARS:
        end = pd.Timestamp(year=y, month=FYE[t], day=1) + pd.offsets.MonthEnd(0)
        if end > SNAP + pd.DateOffset(months=3):
            return y
def trail_year(t, m="rev"):
    ys = [y for y in YEARS if stat(t, m, y) == "A"]
    return ys[-1] if ys else None

# prices (user's adjusted-close file, last date) and FX
pw = openpyxl.load_workbook(PRICE_FILE, data_only=True, read_only=True)["Prices"]
prow = list(pw.iter_rows(values_only=True))
hdr = prow[0]; last = {}
for r in prow[1:]:
    for i, v in enumerate(r[1:], 1):
        if isinstance(v, (int, float)): last[hdr[i]] = (v, r[0])
PRICE_DATE = max(d for _, d in last.values())
quotes = json.load(open(os.path.join(HERE, "fmp_quotes.json")))
EURUSD = last["EURUSD"][0]

def price_local(t):
    p = last[FMP_SYM.get(t, t)][0]
    return p / EURUSD if t == "ASML" else p     # ASML ADR in USD -> EUR to match VA

def mcap_local(t):
    return price_local(t) * quotes[t]["shares"]

def eps_metric(t):
    return "op" if has(t, "op") else "gaap"

# ------------------------------------------------------------------ formatting helpers
def money(v, t, metric, short=False):
    s = SYM[CCY.get(t, "USD")]
    if metric == "rev":
        x = v * REV_SCALE.get(t, 1e6)
        for div, suf in ((1e12, "tn"), (1e9, "bn"), (1e6, "m")):
            if abs(x) >= div:
                q = x / div
                txt = f"{q:,.1f}" if abs(q) < 10 else f"{q:,.0f}"
                return f"{'-' if q < 0 else ''}{s}{txt.lstrip('-')}{suf}"
        return f"{s}{x:,.0f}"
    if abs(v) >= 1000:
        return f"{'-' if v < 0 else ''}{s}{abs(v):,.0f}"
    return f"{'-' if v < 0 else ''}{s}{abs(v):,.2f}"

def unit_label(t, metric):
    s = SYM[CCY.get(t, "USD")] or "KRW "
    if metric != "rev":
        return f"{s.strip()} per share"
    scale = REV_SCALE.get(t, 1e6)
    mx = np.nanmax([val(t, "rev", y) for y in YEARS]) * scale
    return f"{s}tn".replace("KRW tn", "KRW tn") if mx >= 1e12 else f"{s}bn"

def pct(v, signed=True):
    if v is None or np.isnan(v): return "n/m"
    a = abs(v) * 100
    txt = f"{a:,.1f}" if a < 10 else f"{a:,.0f}"
    sign = "+" if v > 0 else "−" if v < 0 else ""
    return (sign if signed else ("−" if v < 0 else "")) + txt + "%"

def growth(a, b):
    if np.isnan(a) or np.isnan(b) or a <= 0: return np.nan
    return b / a - 1

METRIC_NAME = {"rev": "revenue", "gaap": "GAAP EPS", "op": "operating EPS"}
METRIC_TITLE = {"rev": "Revenue", "gaap": "GAAP diluted EPS", "op": "Operating (non-GAAP) diluted EPS"}
SOURCE = f"Source: Visible Alpha ({SNAP:%-d %b %Y}); company reports."
SLUG = {"rev": "revenue", "gaap": "gaap_eps", "op": "operating_eps"}

def fy_label(t, y):
    return f"FY{str(y)[2:]}" + ("" if stat(t, "rev", y) == "A" or stat(t, eps_metric(t), y) == "A" else "E")

def fye_note(t):
    m = FYE[t]
    return "Fiscal year ends in " + MONTH[m] + ("" if m != 12 else " (calendar year)") + "."

saved = []
def save(fig, name):
    p = os.path.join(OUT, name)
    fig.savefig(p, dpi=DPI, facecolor="white")
    plt.close(fig); saved.append(p)
    return p

# ------------------------------------------------------------------ 1) company charts
def company_chart(t, metric, sec_no):
    ys = [y for y in YEARS if not np.isnan(val(t, metric, y))]
    vals = [val(t, metric, y) for y in ys]; sts = [stat(t, metric, y) for y in ys]
    fy = fwd_year(t); name = SHORT[t]; mname = METRIC_NAME[metric]
    if metric == "op" and t == "WELL": mname = "normalized FFO per share"
    # headline: growth in the first consensus year
    first_e = fy if (fy in ys and stat(t, metric, fy) == "E") else next((y for y, s in zip(ys, sts) if s == "E"), None)
    g = growth(val(t, metric, first_e - 1), val(t, metric, first_e)) if first_e and first_e - 1 in ys else np.nan
    if first_e is None:
        title = f"{name} {mname} by fiscal year"
    elif np.isnan(g):
        prev = val(t, metric, first_e - 1)
        title = (f"Consensus expects {name} {mname} to turn positive in FY{str(first_e)[2:]}"
                 if prev <= 0 < val(t, metric, first_e) else f"Consensus {name} {mname} by fiscal year")
    else:
        title = f"Consensus expects {name} {mname} to {'grow' if g >= 0 else 'fall'} {pct(abs(g), signed=False)} in FY{str(first_e)[2:]}"
    hist = con[(con.Ticker == t) & (con.metric == metric) & (con.Fiscal_Year == fy)].dropna(subset=["val"]).sort_values("asof")
    two = len(hist) >= 3
    legend = [("box", "Reported", (REPORTED, None, None)), ("box", "Consensus", (CONSENSUS, EDGE, "--"))]
    if two:
        legend.append(("line", f"FY{str(fy)[2:]} {mname} consensus over time", (BLACK, "-")))
    unit = unit_label(t, metric)
    note = fye_note(t)
    if t == "SPCX": note += " FY25 is Visible Alpha actuals; no report date is listed (pre-IPO)."
    if t == "WELL" and metric == "op": note += " Operating line is normalized FFO per diluted share."
    if t == "KLAC" and metric != "rev": note += " Per-share figures adjusted for the 2026 share split."
    if metric == "op" and t in ("ASML", "2330", "000660"): return None
    fig = fig_base(title, legend, None, SOURCE, "Note: " + note)
    ax = fig.add_axes([0.085, 0.17, 0.43 if two else 0.86, 0.58])
    fig.text(0.05, 0.81, f"{name.upper()} {METRIC_TITLE[metric].upper()} BY FISCAL YEAR, {unit.upper()}",
             fontproperties=F["700"], fontsize=11.5, color=INK, va="center")
    x = np.arange(len(ys))
    scale = REV_SCALE.get(t, 1e6) if metric == "rev" else 1
    div = 1e12 if unit.endswith("tn") else 1e9 if unit.endswith("bn") else 1
    plotv = [v * scale / div for v in vals]
    for xi, v, s in zip(x, plotv, sts):
        if s == "A":
            ax.bar(xi, v, width=0.64, color=REPORTED, linewidth=0, zorder=3)
        else:
            ax.bar(xi, v, width=0.64, color=CONSENSUS, edgecolor=EDGE, linestyle=(0, (3, 2)), linewidth=1.1, zorder=3)
    lo, hi = min(0, min(plotv)), max(0, max(plotv))
    span = hi - lo or 1
    ax.set_ylim(lo - (0.12 * span if lo < 0 else 0), hi + 0.24 * span)
    for i, (xi, v) in enumerate(zip(x, plotv)):
        last_bar = i == len(x) - 1
        off = 0.025 * span
        yv = v + off if v >= 0 else v - off
        ax.text(xi, yv, money(vals[i], t, metric), ha="center", va="bottom" if v >= 0 else "top",
                fontproperties=F["700" if last_bar else "400"], fontsize=13.5, color=INK)
        if i > 0:
            gg = growth(vals[i - 1], vals[i])
            if not np.isnan(gg):
                ax.text(xi, yv + (0.075 * span if v >= 0 else -0.075 * span), pct(gg), ha="center",
                        va="bottom" if v >= 0 else "top", fontproperties=F["400"], fontsize=12.5, color=INK2)
    ax.axhline(0, color="#9a9a9a", linewidth=1, zorder=4)
    ax.set_xticks(x, [f"FY{str(y)[2:]}{'' if s == 'A' else 'E'}" for y, s in zip(ys, sts)])
    s = SYM[CCY.get(t, "USD")]
    suf = "tn" if unit.endswith("tn") else "bn" if unit.endswith("bn") else ""
    ax.yaxis.set_major_locator(plt.MaxNLocator(5))
    ax.yaxis.set_major_formatter(FuncFormatter(
        lambda v, _: (f"{'-' if v < 0 else ''}{s}{abs(v):,.0f}{suf}" if metric == "rev" or abs(v) >= 100 or float(v).is_integer()
                      else f"{'-' if v < 0 else ''}{s}{abs(v):,.1f}")))
    style_axes(ax)
    if two:
        ax2 = fig.add_axes([0.62, 0.17, 0.30, 0.58])
        fig.text(0.575, 0.81, f"FY{str(fy)[2:]} {METRIC_TITLE[metric].upper() if metric != 'rev' else 'REVENUE'} CONSENSUS, {unit.upper()}",
                 fontproperties=F["700"], fontsize=11.5, color=INK, va="center")
        hv = hist["val"].values * scale / div
        ax2.plot(hist["asof"], hv, color=BLACK, linewidth=2.6, marker="o", markersize=5.5, zorder=3)
        lo2, hi2 = hv.min(), hv.max(); pad = (hi2 - lo2) * 0.25 or abs(hi2) * 0.05 or 1
        ax2.set_ylim(lo2 - pad, hi2 + pad)
        ax2.yaxis.set_major_locator(plt.MaxNLocator(5))
        dec = 0 if (metric == "rev" and (hi2 - lo2) > 5) or abs(hi2) >= 100 else 1 if metric == "rev" else 2
        ax2.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f"{'-' if v < 0 else ''}{s}{abs(v):,.{dec}f}{suf}"))
        import matplotlib.dates as mdates
        ax2.xaxis.set_major_locator(mdates.MonthLocator(bymonth=[11, 3, 7]))
        ax2.xaxis.set_major_formatter(mdates.DateFormatter("%b %y"))
        first, lastv = hist["val"].values[0], hist["val"].values[-1]
        ax2.annotate(money(first, t, metric), (hist["asof"].iloc[0], hv[0]), xytext=(0, -20 if hv[0] <= hv.mean() else 12),
                     textcoords="offset points", ha="center", fontproperties=F["400"], fontsize=13, color=INK)
        ax2.annotate(money(lastv, t, metric), (hist["asof"].iloc[-1], hv[-1]), xytext=(10, 0), textcoords="offset points",
                     ha="left", va="center", fontproperties=F["700"], fontsize=13.5, color=INK)
        style_axes(ax2); relabel(ax2)
        ax2.set_xlim(hist["asof"].iloc[0] - pd.Timedelta(days=20), hist["asof"].iloc[-1] + pd.Timedelta(days=45))
    relabel(ax)
    return save(fig, f"{sec_no}_{SECTOR_OF[t].lower().replace('other', 'everything_else')}_company_{t}_{SLUG[metric]}.png")

# ------------------------------------------------------------------ 2) sector indexed
def index_table(sector, metric):
    rows, excluded = {}, {}
    for t in SECTORS[sector]:
        if not has(t, metric):
            excluded[t] = "no line in VA"; continue
        b = val(t, metric, 2025)
        series = np.array([val(t, metric, y) for y in YEARS])
        if np.isnan(b) or b <= 0:
            excluded[t] = "negative FY25 base"; continue
        if (series <= 0).any():
            excluded[t] = "negative in a later year"; continue
        rows[t] = series / b * 100
    return pd.DataFrame(rows, index=YEARS).T, excluded

def sector_index_chart(sector, metric, sec_no):
    idx, excluded = index_table(sector, metric)
    med = idx.median()
    mult = med[2029] / 100
    mname = METRIC_NAME[metric]
    sname = {"Cybersecurity": "cybersecurity", "Tech": "tech", "Other": "everything-else"}[sector]
    title = (f"Consensus has median {sname} {mname} at {mult:,.1f}x its FY25 level by FY29" if sector != "Other" else
             f"Everything else: consensus has median {mname} at {mult:,.1f}x FY25 by FY29")
    if sector == "Cybersecurity": title = title.replace("cybersecurity", "cybersecurity")
    notes = ["Each company rebased to its own FY25 = 100, so currencies drop out. Median of the company indices."]
    if excluded:
        notes.append("Excluded: " + ", ".join(f"{SHORT[t]} ({w})" for t, w in excluded.items()) + ".")
    if metric == "op" and "WELL" in idx.index: notes.append("Welltower uses normalized FFO per share.")
    use_log = idx.values.max() / max(idx.values.min(), 1) > 8
    legend = [("line", "Sector median", (TEAL_DARK, "-")), ("line", "Company", (TEAL_LIGHT, "-"))]
    fig = fig_base(title, legend, f"{'Everything else' if sector == 'Other' else sector} {METRIC_TITLE[metric]}, indexed FY25 = 100" + (" (log scale)" if use_log else ""),
                   SOURCE, "Note: " + " ".join(notes) + " Fiscal years follow each company's own calendar.")
    ax = fig.add_axes([0.085, 0.17, 0.74, 0.58])
    x = np.arange(len(YEARS))
    ax.axvspan(1.5, len(YEARS) - 0.5 + 0.15, color=SHADE, zorder=0, linewidth=0)
    ax.text((1.5 + len(YEARS) - 0.35) / 2, 1.0, "Forecast (consensus)", transform=ax.get_xaxis_transform(), ha="center",
            va="bottom", fontproperties=F["400"], fontsize=13, color=INK2)
    for t in idx.index:
        ax.plot(x, idx.loc[t].values, color=TEAL_LIGHT, linewidth=1.8, zorder=2)
    ax.plot(x, med.values, color=TEAL_DARK, linewidth=3.4, zorder=5, marker="o", markersize=7)
    for xi, v in zip(x, med.values):
        if xi == 0 or xi == len(x) - 1: continue
        ax.annotate(f"{v:,.0f}", (xi, v), xytext=(0, 12), textcoords="offset points", ha="center",
                    fontproperties=F["700"], fontsize=13, color=INK, zorder=6)
    if use_log:
        ax.set_yscale("log")
        lo, hi = min(idx.values.min(), 100) * 0.8, idx.values.max() * 1.3
        ax.set_ylim(lo, hi)
        ax.set_yticks([v for v in [25, 50, 100, 200, 500, 1000, 2000, 5000, 10000] if lo <= v <= hi])
        ax.yaxis.set_minor_locator(NullLocator())
    else:
        lo, hi = min(idx.values.min(), 100), idx.values.max()
        ax.set_ylim(lo - (hi - lo) * 0.08, hi + (hi - lo) * 0.12)
        ax.yaxis.set_major_locator(plt.MaxNLocator(6))
    ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f"{v:,.0f}"))
    ax.axhline(100, color="#9a9a9a", linewidth=1, linestyle=(0, (4, 3)), zorder=1)
    ax.set_xticks(x, ["FY25", "FY26", "FY27E", "FY28E", "FY29E"])
    ax.set_xlim(-0.25, len(YEARS) - 1 + 0.15)
    style_axes(ax)
    # end labels with collision nudging
    fig.canvas.draw()
    tr = ax.transData
    pts = sorted([(tr.transform((x[-1], idx.loc[t, 2029]))[1], t) for t in idx.index] +
                 [(tr.transform((x[-1], med[2029]))[1], "__med")])
    placed = []
    for ypix, t in pts:
        yp = max(ypix, placed[-1] + 30) if placed else ypix
        placed.append(yp)
    # shift down if overflow
    top = tr.transform((0, ax.get_ylim()[1]))[1] + 20
    if placed and placed[-1] > top:
        shift = placed[-1] - top; placed = [p - shift for p in placed]
    for (ypix, t), yp in zip(pts, placed):
        yd = tr.inverted().transform((0, yp))[1]
        v = med[2029] if t == "__med" else idx.loc[t, 2029]
        label = f"Median {v:,.0f}" if t == "__med" else f"{SHORT[t]} {v:,.0f}"
        ax.annotate(label, (x[-1], v), xytext=(x[-1] + 0.12, yd), textcoords="data", va="center", annotation_clip=False,
                    fontproperties=F["700" if t == "__med" else "400"], fontsize=12.5 if t == "__med" else 11.5,
                    color=INK if t == "__med" else INK2,
                    arrowprops=dict(arrowstyle="-", color="#c8c8c8", linewidth=0.8, shrinkA=0, shrinkB=2) if abs(yp - ypix) > 3 else None)
    relabel(ax)
    return save(fig, f"{sec_no}_{sector.lower().replace('other', 'everything_else')}_indexed_{SLUG[metric]}.png"), idx, med, excluded

# ------------------------------------------------------------------ 3) valuation & growth ranked bars
def valuation_rows(sector):
    out = []
    for t in SECTORS[sector]:
        fy = fwd_year(t); ty = trail_year(t, "rev")
        em = eps_metric(t); tye = trail_year(t, em)
        p = price_local(t); mc = mcap_local(t)
        rev_t, rev_f = val(t, "rev", ty), val(t, "rev", fy)
        eps_t, eps_f = val(t, em, tye), val(t, em, fy)
        sc = REV_SCALE.get(t, 1e6)
        out.append(dict(
            Ticker=t, Company=SHORT[t], Sector=sector, Currency=CCY.get(t, "USD"), Price_local=p, Price_date=PRICE_DATE,
            Shares=quotes[t]["shares"], Market_cap_local=mc, Trailing_FY=ty, Forward_FY=fy, EPS_line=METRIC_NAME[em],
            Revenue_trailing=rev_t, Revenue_forward=rev_f, EPS_trailing=eps_t, EPS_forward=eps_f,
            PS_trailing=mc / (rev_t * sc) if rev_t > 0 else np.nan,
            PS_forward=mc / (rev_f * sc) if rev_f > 0 else np.nan,
            PE_trailing=p / eps_t if eps_t > 0 else np.nan,
            PE_forward=p / eps_f if eps_f > 0 else np.nan,
            Sales_growth_fwd=growth(val(t, "rev", fy - 1), rev_f),
            EPS_growth_fwd=growth(val(t, em, fy - 1), eps_f),
        ))
    return pd.DataFrame(out)

SHORTM = {"PS_trailing": "trailing P/S", "PS_forward": "forward P/S", "PE_trailing": "trailing P/E", "PE_forward": "forward P/E", "Sales_growth_fwd": "forward sales growth", "EPS_growth_fwd": "forward EPS growth"}
VAL_METRICS = [
    ("PS_trailing", "Price to sales, trailing", "x", "price ÷ last reported fiscal-year revenue per share"),
    ("PS_forward", "Price to forward sales", "x", "price ÷ forward fiscal-year consensus revenue per share"),
    ("PE_trailing", "Price to earnings, trailing", "x", "price ÷ last reported fiscal-year diluted EPS"),
    ("PE_forward", "Price to forward earnings", "x", "price ÷ forward fiscal-year consensus diluted EPS"),
    ("Sales_growth_fwd", "Sales growth, forward fiscal year", "%", "forward fiscal-year consensus revenue vs the prior fiscal year"),
    ("EPS_growth_fwd", "EPS growth, forward fiscal year", "%", "forward fiscal-year consensus diluted EPS vs the prior fiscal year"),
]

def fmt_metric(v, kind):
    if np.isnan(v): return "n/m"
    return f"{v:,.1f}x" if kind == "x" and v < 100 else f"{v:,.0f}x" if kind == "x" else pct(v)

def ranked_chart(df, col, label, kind, defn, sector, sec_no, i):
    disp = "Everything else" if sector == "Other" else sector
    d = df[["Company", "Ticker", col, "Forward_FY", "Trailing_FY"]].copy()
    good = d[col].notna()
    d = pd.concat([d[good].sort_values(col, ascending=False), d[~good]])
    med = d[col].median()
    top = d.iloc[0]
    if kind == "x":
        title = f"{top.Company} trades on {fmt_metric(top[col], kind)} {label.lower().replace('price to ', '').replace(', trailing', ' (trailing)')}, the highest in {sector.lower() if sector != 'Other' else 'the other group'}"
        title = f"{disp}: median {SHORTM[col]} is {fmt_metric(med, kind)}; {top.Company} is highest at {fmt_metric(top[col], kind)}"
    else:
        title = f"{disp}: median {SHORTM[col]} is {pct(med)}; {top.Company} leads at {pct(top[col])}"

    legend = [("box", "Company", (TEAL_LIGHT, None, None)), ("line", f"{disp} median", ("#6b6b6b", (0, (4, 3))))]
    em_note = "EPS uses operating (non-GAAP) diluted EPS; GAAP where VA has no operating line (ASML, TSMC, SK Hynix). " if "EPS" in label or "earnings" in label else ""
    fy_note = "Forward FY = first fiscal year ending more than three months after the snapshot (FY28 for January year-ends, FY27 for the rest). "
    fig = fig_base(title, legend, f"{label}, as at {PRICE_DATE:%-d %b %Y}",
                   f"Source: Visible Alpha ({SNAP:%-d %b %Y}); price data ({PRICE_DATE:%-d %b %Y} close); FMP shares outstanding.",
                   "Note: " + label + " = " + defn + ". " + (fy_note if "forward" in label.lower() or "growth" in label.lower() else "") +
                   em_note + "n/m = negative or not meaningful. Local currency throughout; ASML price converted to EUR.")
    n = len(d)
    ax = fig.add_axes([0.2, 0.15, 0.7, 0.6])
    y = np.arange(n)[::-1]
    vals = d[col].values.astype(float)
    plotv = np.where(np.isnan(vals), 0, vals * (100 if kind == "%" else 1))
    hgt = 0.62 if n > 6 else 0.5
    ax.barh(y, plotv, height=hgt, color=TEAL_LIGHT, linewidth=0, zorder=3)
    mx = np.nanmax(np.abs(plotv)) or 1
    lo = min(0, np.nanmin(plotv)); hi = max(0, np.nanmax(plotv))
    ax.set_xlim(lo - (0.12 * (hi - lo) if lo < 0 else 0), hi + 0.14 * (hi - lo))
    for yi, v, raw in zip(y, plotv, vals):
        txt = fmt_metric(raw, kind)
        ax.text(v + (0.01 * (hi - lo) if v >= 0 else -0.01 * (hi - lo)), yi, txt, va="center",
                ha="left" if v >= 0 else "right", fontproperties=F["400"], fontsize=12.5, zorder=6,
                color=INK if not np.isnan(raw) else MUTED,
                bbox=dict(boxstyle="square,pad=0.15", facecolor="white", edgecolor="none"))
    mv = med * (100 if kind == "%" else 1)
    ax.axvline(mv, color="#6b6b6b", linewidth=2, linestyle=(0, (4, 3)), zorder=4)
    ax.text(mv, n - 0.35, f"Median {fmt_metric(med, kind)}", ha="center", va="bottom", zorder=6,
            fontproperties=F["700"], fontsize=12.5, color=INK,
            bbox=dict(boxstyle="square,pad=0.15", facecolor="white", edgecolor="none"))
    ax.axvline(0, color="#9a9a9a", linewidth=1, zorder=4)
    ax.set_yticks(y, d.Company.values)
    ax.set_ylim(-0.7, n - 0.1)
    ax.xaxis.set_major_locator(plt.MaxNLocator(6))
    ax.xaxis.set_major_formatter(FuncFormatter(lambda v, _: f"{v:,.0f}x" if kind == "x" else f"{'+' if v > 0 else '−' if v < 0 else ''}{abs(v):,.0f}%"))
    style_axes(ax, ygrid=False, xgrid=True)
    relabel(ax)
    for lab in ax.get_yticklabels():
        lab.set_fontsize(13.5); lab.set_color(INK)
    return save(fig, f"{sec_no}_{sector.lower().replace('other', 'everything_else')}_{i:02d}_{col.lower()}.png")

# ------------------------------------------------------------------ run
if __name__ == "__main__":
    allval, idx_tables = [], []
    for si, sector in enumerate(SECTORS, 1):
        df = valuation_rows(sector); allval.append(df)
        for i, (col, label, kind, defn) in enumerate(VAL_METRICS, 1):
            ranked_chart(df, col, label, kind, defn, sector, f"{si}a", i)
        for mi, metric in enumerate(["rev", "gaap", "op"]):
            p, idx, med, exc = sector_index_chart(sector, metric, f"{si}{'bcd'[mi]}")
            t = idx.copy(); t.loc["Median"] = med; t.insert(0, "Metric", METRIC_TITLE[metric]); t.insert(0, "Sector", sector)
            idx_tables.append(t.reset_index().rename(columns={"index": "Ticker"}))
            for tk in SECTORS[sector]:
                if has(tk, metric):
                    company_chart(tk, metric, f"{si}{'bcd'[mi]}")
    vals = pd.concat(allval)
    long = A[["Ticker", "Company_Name", "metric", "Line_Item", "Currency", "Fiscal_Year", "value", "status"]].copy()
    long["Sector"] = long.Ticker.map(SECTOR_OF)
    long["metric"] = long.metric.map(METRIC_TITLE)
    with pd.ExcelWriter(os.path.join(OUT, "chart_data.xlsx")) as xw:
        vals.to_excel(xw, sheet_name="Valuation & growth", index=False)
        pd.concat(idx_tables).round(1).to_excel(xw, sheet_name="Indexed FY25=100", index=False)
        long.rename(columns={"status": "A=reported E=consensus"}).sort_values(["Sector", "Ticker", "metric", "Fiscal_Year"]).to_excel(xw, sheet_name="FY values", index=False)
    print(len(saved), "charts")
    print(vals[["Ticker", "Trailing_FY", "Forward_FY", "PS_trailing", "PS_forward", "PE_trailing", "PE_forward", "Sales_growth_fwd", "EPS_growth_fwd"]].round(2).to_string())
