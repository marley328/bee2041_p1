"""All-company ranked bar charts for P/E and P/S (trailing and forward), coloured by sector."""
import os, sys
OUTDIR = sys.argv[1] if len(sys.argv) > 1 else "/tmp/valbars"
sys.argv = [sys.argv[0], "/tmp/_unused_build_out"]
sys.path.insert(0, os.path.dirname(__file__))
import build as B
from house import *
import numpy as np, pandas as pd
from matplotlib.ticker import FuncFormatter

os.makedirs(OUTDIR, exist_ok=True)
SEC_ORDER = ["Tech", "Cybersecurity", "Other"]
SEC_NAME = {"Tech": "Tech", "Cybersecurity": "Cybersecurity", "Other": "Everything else"}
SEC_COLOR = {"Tech": TEAL_DARK, "Cybersecurity": SALMON, "Other": TEAL_LIGHT}
LINE_STYLE = {"Tech": (0, (5, 3)), "Cybersecurity": (0, (2, 2)), "Other": (0, (8, 3, 2, 3))}
df = pd.concat([B.valuation_rows(s) for s in SEC_ORDER], ignore_index=True)
SRC = (f"Source: Visible Alpha ({B.SNAP:%-d %b %Y}); price data ({B.PRICE_DATE:%-d %b %Y} close); FMP shares outstanding.")
FWD = "Forward = FY28 for January year-ends, FY27 for the rest; trailing = last reported fiscal year."

SPECS = [
    ("PE_forward", "forward P/E", "Price to forward earnings",
     "P/E = share price ÷ consensus diluted EPS: operating (non-GAAP) where Visible Alpha has it, GAAP for ASML, TSMC and "
     "SK Hynix; Welltower uses normalized FFO per share. "),
    ("PE_trailing", "trailing P/E", "Price to trailing earnings",
     "P/E = share price ÷ last reported diluted EPS: operating where Visible Alpha has it, GAAP for ASML, TSMC and SK Hynix; "
     "Welltower uses normalized FFO per share. n/m = negative or no reported EPS. "),
    ("PS_forward", "forward P/S", "Price to forward sales", "P/S = market value ÷ consensus revenue, local currency. "),
    ("PS_trailing", "trailing P/S", "Price to trailing sales", "P/S = market value ÷ last reported revenue, local currency. "),
]

def chart(col, short, label, defn, fname):
    d = df[["Company", "Sector", col]].copy()
    good = d[col].notna()
    d = pd.concat([d[good].sort_values(col, ascending=False), d[~good]]).reset_index(drop=True)
    med = {s: df[df.Sector == s][col].median() for s in SEC_ORDER}
    top, low = d[good.loc[d.index] if False else d[col].notna()].iloc[0], d[d[col].notna()].iloc[-1]
    title = (f"{top.Company} is the most expensive on {short} at {top[col]:,.0f}x; "
             f"{low.Company} the cheapest at {low[col]:,.1f}x")
    legend = [("box", f"{SEC_NAME[s]} (median {med[s]:,.1f}x)", (SEC_COLOR[s], None, None)) for s in SEC_ORDER]
    fig = fig_base(title, legend, f"{label}, all 23 companies, as at {B.PRICE_DATE:%-d %b %Y}", SRC,
                   "Note: " + defn + FWD + " Sector medians shown in the key.")
    n = len(d)
    ax = fig.add_axes([0.2, 0.155, 0.72, 0.6])
    y = np.arange(n)[::-1]
    vals = d[col].values.astype(float)
    hi = np.nanmax(vals)
    ax.barh(y, np.nan_to_num(vals), height=0.68, color=[SEC_COLOR[s] for s in d.Sector], linewidth=0, zorder=3)
    for yi, v in zip(y, vals):
        txt = "n/m" if np.isnan(v) else (f"{v:,.1f}x" if v < 100 else f"{v:,.0f}x")
        ax.text((0 if np.isnan(v) else v) + hi * 0.008, yi, txt, va="center", ha="left", fontproperties=F["400"],
                fontsize=11.5, color=MUTED if np.isnan(v) else INK, zorder=6)
    ax.set_yticks(y, d.Company.values)
    ax.set_ylim(-0.7, n - 0.3)
    ax.set_xlim(0, hi * 1.1)
    ax.xaxis.set_major_locator(plt.MaxNLocator(7))
    ax.xaxis.set_major_formatter(FuncFormatter(lambda v, _: f"{v:,.0f}x"))
    style_axes(ax, ygrid=False, xgrid=True); relabel(ax)
    for lab in ax.get_yticklabels():
        lab.set_fontsize(12); lab.set_color(INK)
    p = os.path.join(OUTDIR, fname)
    fig.savefig(p, dpi=DPI, facecolor="white"); plt.close(fig)
    return p

if __name__ == "__main__":
    for i, (col, short, label, defn) in enumerate(SPECS, 1):
        print(chart(col, short, label, defn, f"V{i}_{col.lower()}_all_companies.png"))
    print({s: {c: round(df[df.Sector == s][c].median(), 1) for c, *_ in SPECS} for s in SEC_ORDER})
