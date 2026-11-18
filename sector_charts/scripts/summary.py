"""Four all-company summary charts in house style."""
import os, sys
sys.argv = [sys.argv[0], sys.argv[1] if len(sys.argv) > 1 else "/tmp/summary_out"]
OUTDIR = sys.argv[1]
sys.path.insert(0, os.path.dirname(__file__))
import build as B
from house import *
import numpy as np, pandas as pd
from matplotlib.ticker import FuncFormatter, NullLocator, FixedLocator
from matplotlib.lines import Line2D

os.makedirs(OUTDIR, exist_ok=True)
SEC_ORDER = ["Tech", "Cybersecurity", "Other"]
SEC_NAME = {"Tech": "Tech", "Cybersecurity": "Cybersecurity", "Other": "Everything else"}
SEC_COLOR = {"Tech": TEAL_DARK, "Cybersecurity": SALMON, "Other": BLACK}
df = pd.concat([B.valuation_rows(s) for s in SEC_ORDER], ignore_index=True)
SRC_VAL = (f"Source: Visible Alpha ({B.SNAP:%-d %b %Y}); price data ({B.PRICE_DATE:%-d %b %Y} close); "
           "FMP shares outstanding.")
FWD_NOTE = ("Forward = first fiscal year ending more than three months after the snapshot (FY28 for January "
            "year-ends, FY27 for the rest); trailing = last reported fiscal year.")
saved = []

def rows_layout(order_col):
    """Row positions top-to-bottom, sectors separated by a header row."""
    rows, y, heads = [], 0, []
    for s in SEC_ORDER:
        d = df[df.Sector == s].copy()
        d = pd.concat([d[d[order_col].notna()].sort_values(order_col, ascending=False), d[d[order_col].isna()]])
        heads.append((y, s)); y += 1
        for _, r in d.iterrows():
            rows.append((y, r)); y += 1
        y += 0.35
    return rows, heads, y

def dumbbell(a, b, label_a, label_b, title, sub, kind, fname, note, clip=None, log=False):
    rows, heads, n = rows_layout(b)
    legend = [("box", label_a, ("white", "#7a7a7a", "-")), ("box", label_b, (TEAL_DARK, None, None))]
    fig = fig_base(title, None, sub, SRC_VAL, "Note: " + note)
    # custom legend: circles
    x = 0.05
    for lab, fc, ec in ((label_a, "white", "#6f6f6f"), (label_b, TEAL_DARK, TEAL_DARK)):
        fig.add_artist(Line2D([x + 0.006], [0.857], marker="o", markersize=10, markerfacecolor=fc,
                              markeredgecolor=ec, markeredgewidth=1.8, transform=fig.transFigure, linestyle=""))
        t = fig.text(x + 0.018, 0.857, lab, fontproperties=F["400"], fontsize=13, color=INK2, va="center")
        fig.canvas.draw(); x = t.get_window_extent().transformed(fig.transFigure.inverted()).x1 + 0.03
    fig.add_artist(Line2D([x, x + 0.026], [0.857, 0.857], color="#6b6b6b", linewidth=2, linestyle=(0, (4, 3)),
                          transform=fig.transFigure))
    fig.text(x + 0.032, 0.857, f"Sector median ({label_b})", fontproperties=F["400"], fontsize=13, color=INK2, va="center")
    ax = fig.add_axes([0.22, 0.165, 0.71, 0.605])
    yr = lambda y: n - y
    conv = (lambda v: v * 100) if kind == "%" else (lambda v: v)
    allv = np.concatenate([df[a].dropna().map(conv).values, df[b].dropna().map(conv).values])
    hi = clip if clip else allv.max() * (1.15 if not log else 1.6)
    lo = (allv[allv > 0].min() * 0.7) if log else min(0, allv.min() * 1.1)
    for y, r in rows:
        va, vb = conv(r[a]), conv(r[b])
        yy = yr(y)
        pts = [v for v in (va, vb) if pd.notna(v)]
        if len(pts) == 2:
            ax.plot([min(min(pts), hi), min(max(pts), hi)], [yy, yy], color="#c9c9c9", linewidth=2.2, zorder=2)
        for v, fc, ec in ((va, "white", "#6f6f6f"), (vb, TEAL_DARK, TEAL_DARK)):
            if pd.isna(v):
                continue
            vv = min(v, hi)
            ax.plot([vv], [yy], marker="o", markersize=9, markerfacecolor=fc, markeredgecolor=ec,
                    markeredgewidth=1.8, zorder=4, linestyle="")
            pass
        if any(pd.notna(v) and v > hi for v in (va, vb)):
            ax.annotate(f"{label_a} {B.fmt_metric(r[a], kind)}, {label_b} {B.fmt_metric(r[b], kind)} (off scale)",
                        (hi, yy), xytext=(-14, 0), textcoords="offset points", ha="right", va="center",
                        fontproperties=F["700"], fontsize=11, color=INK, zorder=6,
                        bbox=dict(boxstyle="square,pad=0.15", facecolor="white", edgecolor="none"))
        # forward value label at the right of the forward dot
        if pd.notna(vb) and vb <= hi:
            side = 1 if (pd.isna(va) or vb >= va) else -1
            ax.annotate(B.fmt_metric(r[b], kind), (vb, yy), xytext=(10 * side, 0), textcoords="offset points",
                        ha="left" if side > 0 else "right", va="center", fontproperties=F["700"], fontsize=11, color=INK,
                        bbox=dict(boxstyle="square,pad=0.1", facecolor="white", edgecolor="none"), zorder=5)
        if pd.isna(vb) and pd.isna(va):
            ax.text(lo, yy, " n/m", va="center", fontproperties=F["400"], fontsize=11, color=MUTED)
        elif pd.isna(va):
            pass
    # sector medians (forward), drawn per sector block
    for (hy, s), nxt in zip(heads, heads[1:] + [(n, None)]):
        med = conv(df[df.Sector == s][b].median())
        top, bot = yr(hy) - 0.5, yr(nxt[0]) + 0.5 + (0.35 if nxt[1] else 0.35)
        ax.plot([med, med], [bot, top], color="#6b6b6b", linewidth=1.8, linestyle=(0, (4, 3)), zorder=3)
        ax.text(lo if not log else lo, yr(hy), "", va="center")
        fig_y = ax.transData.transform((0, yr(hy)))[1]
        ax.annotate(f"{SEC_NAME[s].upper()}  ·  median {B.fmt_metric(df[df.Sector == s][b].median(), kind)}",
                    xy=(0, yr(hy)), xycoords=("axes fraction", "data"), xytext=(-0.29, yr(hy)),
                    textcoords=("axes fraction", "data"), va="center", ha="left", fontproperties=F["700"],
                    fontsize=12, color=SEC_COLOR[s])
    ax.set_yticks([yr(y) for y, _ in rows], [r.Company for _, r in rows])
    ax.set_ylim(0.3, n + 0.2)
    if log:
        ax.set_xscale("log")
        ticks = [t for t in [1, 2, 5, 10, 20, 50, 100, 200, 500] if lo <= t <= hi]
        ax.xaxis.set_major_locator(FixedLocator(ticks)); ax.xaxis.set_minor_locator(NullLocator())
    ax.set_xlim(lo, hi * (1.0 if clip else 1.0))
    ax.xaxis.set_major_formatter(FuncFormatter(
        lambda v, _: f"{v:,.0f}x" if kind == "x" else f"{'+' if v > 0 else '−' if v < 0 else ''}{abs(v):,.0f}%"))
    if kind == "%":
        ax.axvline(0, color="#9a9a9a", linewidth=1, zorder=1)
    style_axes(ax, ygrid=False, xgrid=True); relabel(ax)
    for lab in ax.get_yticklabels():
        lab.set_fontsize(12); lab.set_color(INK)
    p = os.path.join(OUTDIR, fname); fig.savefig(p, dpi=DPI, facecolor="white"); plt.close(fig); saved.append(p)

def headline_val(col, kind, verb):
    top = df.loc[df[col].idxmax()]; med = df[col].median()
    return top, med

# 1) P/S
t, med = headline_val("PS_forward", "x", "")
lowest = df.loc[df.PS_forward.idxmin()]
dumbbell("PS_trailing", "PS_forward", "Trailing P/S", "Forward P/S",
         f"{t.Company} trades on {t.PS_forward:,.0f}x forward sales; {lowest.Company} on {lowest.PS_forward:,.1f}x",
         f"Price to sales, all 23 companies, as at {B.PRICE_DATE:%-d %b %Y} (log scale)", "x", "S1_price_to_sales.png",
         "P/S = market value ÷ fiscal-year revenue, local currency. " + FWD_NOTE, log=True)
# 2) P/E
t = df.loc[df.PE_forward.idxmax()]; lo_ = df.loc[df.PE_forward.idxmin()]
dumbbell("PE_trailing", "PE_forward", "Trailing P/E", "Forward P/E",
         f"Forward P/E runs from {lo_.PE_forward:,.1f}x at {lo_.Company} to {t.PE_forward:,.0f}x at {t.Company}",
         f"Price to earnings, all 23 companies, as at {B.PRICE_DATE:%-d %b %Y} (log scale)", "x", "S2_price_to_earnings.png",
         "P/E = share price ÷ diluted EPS: operating (non-GAAP) where Visible Alpha has it, GAAP for ASML, TSMC and SK Hynix; "
         "Welltower uses normalized FFO per share. Missing dot = negative EPS (n/m). " + FWD_NOTE, log=True)
# 3) growth
t = df.loc[df.Sales_growth_fwd.idxmax()]
dumbbell("Sales_growth_fwd", "EPS_growth_fwd", "Sales growth", "EPS growth",
         f"Tech leads forward growth: median sales {df[df.Sector=='Tech'].Sales_growth_fwd.median()*100:,.0f}%, EPS {df[df.Sector=='Tech'].EPS_growth_fwd.median()*100:,.0f}%",
         "Consensus growth in the forward fiscal year vs the prior year, all 23 companies", "%", "S3_sales_and_eps_growth.png",
         "EPS growth uses operating diluted EPS where available (GAAP for ASML, TSMC, SK Hynix). Axis capped at +150%; "
         "values beyond are labelled. " + FWD_NOTE, clip=150)

# 4) indexed: three panels, sector medians
def indexed_panels():
    meds = {}
    for m in ["rev", "gaap", "op"]:
        for s in SEC_ORDER:
            idx, exc = B.index_table(s, m)
            meds[(m, s)] = (idx.median(), len(idx), exc)
    best = max(SEC_ORDER, key=lambda s: meds[("rev", s)][0][2029])
    title = (f"{SEC_NAME[best]} leads: consensus has its median revenue at "
             f"{meds[('rev', best)][0][2029] / 100:,.1f}x FY25 by FY29")
    legend = [("line", SEC_NAME[s], (SEC_COLOR[s], "-")) for s in SEC_ORDER]
    excl = []
    for (m, s), (_, k, exc) in meds.items():
        for tk, why in exc.items():
            if why != "no line in VA":
                excl.append(f"{B.SHORT[tk]} ({'GAAP' if m == 'gaap' else 'operating'} EPS)")
    fig = fig_base(title, legend, None, f"Source: Visible Alpha ({B.SNAP:%-d %b %Y}); company reports.",
                   "Note: Cybersecurity is left off the GAAP EPS panel: only Palo Alto has a positive FY25 base. Each company rebased to its own FY25 = 100; line = median of the companies in the group. "
                   "FY27–FY29 are consensus; FY26 mixes reported and consensus. Excluded for a negative base: "
                   + ", ".join(sorted(set(excl))) + ". Operating EPS has no line for ASML, TSMC, SK Hynix; Welltower uses "
                   "normalized FFO per share. Fiscal years follow each company's calendar.")
    titles = {"rev": "REVENUE", "gaap": "GAAP DILUTED EPS", "op": "OPERATING DILUTED EPS"}
    ymax = max(meds[k][0].max() for k in meds) * 1.12
    x = np.arange(5)
    for i, m in enumerate(["rev", "gaap", "op"]):
        ax = fig.add_axes([0.07 + i * 0.31, 0.21, 0.25, 0.53])
        fig.text(0.05 + i * 0.31, 0.81, titles[m] + ", FY25 = 100", fontproperties=F["700"], fontsize=11.5, color=INK, va="center")
        ax.axvspan(1.5, 4.4, color=SHADE, zorder=0, linewidth=0)
        ax.axhline(100, color="#9a9a9a", linewidth=1, linestyle=(0, (4, 3)), zorder=1)
        ends = []
        for s in SEC_ORDER:
            med, k, _ = meds[(m, s)]
            if k < 2:
                continue
            ax.plot(x, med.values, color=SEC_COLOR[s], linewidth=3, marker="o", markersize=5.5, zorder=4)
            ends.append((med[2029], s))
        # end labels, nudged
        fig.canvas.draw(); tr = ax.transData; placed = []
        for v, s in sorted(ends):
            yp = tr.transform((4, v))[1]
            if placed and yp - placed[-1] < 24: yp = placed[-1] + 24
            placed.append(yp)
            ax.annotate(f"{v:,.0f}", (4, v), xytext=(4.15, tr.inverted().transform((0, yp))[1]), textcoords="data",
                        va="center", fontproperties=F["700"], fontsize=12.5, color=SEC_COLOR[s], annotation_clip=False)
        ax.set_ylim(60, ymax)
        ax.set_xlim(-0.3, 4.4)
        ax.set_xticks(x, ["FY25", "FY26", "FY27E", "FY28E", "FY29E"])
        ax.yaxis.set_major_locator(plt.MaxNLocator(6))
        ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f"{v:,.0f}"))
        style_axes(ax); relabel(ax)
        for lab in ax.get_xticklabels(): lab.set_fontsize(11.5)
        ax.text(2.95, 1.0, "Consensus", transform=ax.get_xaxis_transform(), ha="center", va="bottom",
                fontproperties=F["400"], fontsize=11.5, color=INK2)
    p = os.path.join(OUTDIR, "S4_sector_indexed_revenue_and_eps.png")
    fig.savefig(p, dpi=DPI, facecolor="white"); plt.close(fig); saved.append(p)
    return meds

meds = indexed_panels()
print("\n".join(saved))
for (m, s), (med, k, exc) in meds.items():
    print(m, s, k, med.round(0).tolist())
