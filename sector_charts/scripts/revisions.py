"""Consensus revision charts: heatmap (all companies x metrics x windows) and sector revision paths."""
import os, sys
OUTDIR = sys.argv[1] if len(sys.argv) > 1 else "/tmp/rev_out"
sys.argv = [sys.argv[0], "/tmp/_unused_build_out"]
sys.path.insert(0, os.path.dirname(__file__))
import build as B
from house import *
import numpy as np, pandas as pd
from matplotlib.colors import to_rgb
from matplotlib.patches import Rectangle
from matplotlib.ticker import FuncFormatter
import matplotlib.dates as mdates

os.makedirs(OUTDIR, exist_ok=True)
SEC_ORDER = ["Tech", "Cybersecurity", "Other"]
SEC_NAME = {"Tech": "Tech", "Cybersecurity": "Cybersecurity", "Other": "Everything else"}
SEC_COLOR = {"Tech": TEAL_DARK, "Cybersecurity": SALMON, "Other": BLACK}
METRICS = ["rev", "gaap", "op"]
MLABEL = {"rev": "REVENUE", "gaap": "GAAP EPS", "op": "OPERATING EPS"}
WINDOWS = [(-30, "1M"), (-90, "3M"), (-365, "12M")]
LOOKBACKS = [-365, -180, -120, -90, -60, -30, -14, -7, 0]
CAP = 0.5
con = B.con

def cons(t, m, fy, lb):
    r = con[(con.Ticker == t) & (con.metric == m) & (con.Fiscal_Year == fy) & (con.Lookback_Days == lb)]
    return float(r["val"].iloc[0]) if len(r) and pd.notna(r["val"].iloc[0]) else np.nan

def revision(t, m, lb):
    fy = B.fwd_year(t)
    now, then = cons(t, m, fy, 0), cons(t, m, fy, lb)
    if np.isnan(now) or np.isnan(then) or then <= 0:
        return np.nan, ("–" if not B.has(t, m) else "n/m")
    return now / then - 1, None

# ---------------------------------------------------------------- data table
rows = []
for s in SEC_ORDER:
    for t in B.SECTORS[s]:
        rec = {"Sector": s, "Ticker": t, "Company": B.SHORT[t], "Forward_FY": B.fwd_year(t)}
        for m in METRICS:
            for lb, w in WINDOWS:
                v, flag = revision(t, m, lb)
                rec[f"{m}_{w}"] = v; rec[f"{m}_{w}_flag"] = flag
        rows.append(rec)
R = pd.DataFrame(rows)
COLS = [f"{m}_{w}" for m in METRICS for _, w in WINDOWS]

# ---------------------------------------------------------------- 1) heatmap
NEG, POS, MID = to_rgb(SALMON), to_rgb(TEAL_DARK), to_rgb("#f0efec")
def cell_color(v):
    if np.isnan(v):
        return "#ffffff"
    x = max(-1, min(1, v / CAP))
    a = abs(x) ** 0.8
    end = POS if x > 0 else NEG
    return tuple(MID[i] + (end[i] - MID[i]) * a for i in range(3))

def text_color(v):
    return "white" if (not np.isnan(v) and abs(v) / CAP > 0.62) else INK

def fmt(v):
    a = abs(v) * 100
    s = f"{a:,.1f}" if a < 10 else f"{a:,.0f}"
    return ("+" if v > 0 else "−" if v < 0 else "") + s + "%"

def heatmap():
    ordered = []
    for s in SEC_ORDER:
        d = R[R.Sector == s].copy()
        d["_k"] = d["op_12M"].fillna(d["gaap_12M"])
        d = pd.concat([d[d._k.notna()].sort_values("_k", ascending=False), d[d._k.isna()]])
        ordered.append(("median", s, d))
    tech_med = R[R.Sector == "Tech"]["op_12M"].median()
    cyb_med = R[R.Sector == "Cybersecurity"]["op_12M"].median()
    oth_med = R[R.Sector == "Other"]["op_12M"].median()
    title = (f"Tech operating EPS consensus is up {fmt(tech_med).lstrip('+')} in a year, "
             f"versus {fmt(cyb_med).lstrip('+')} for cybersecurity")
    fig = fig_base(title, None,
                   "Change in forward fiscal-year consensus over 1, 3 and 12 months, to 18 Nov 2026",
                   f"Source: Visible Alpha ({B.SNAP:%-d %b %Y}).",
                   "Note: Forward FY = FY28 for January year-ends, FY27 for the rest. Colour scale capped at ±50%. "
                   "n/m = no consensus, or a negative one, at the start of the window (SpaceX has no 12M history); – = no such line in Visible Alpha "
                   "(operating EPS for ASML, TSMC, SK Hynix). Welltower operating = normalized FFO per share. The 12M "
                   "change can include brokers starting coverage. Sectors sorted by 12M operating EPS revision.")
    # legend: diverging swatches
    lx = 0.05
    for v, lab in ((-0.5, "Cut"), (0, "No change"), (0.5, "Upgrade")):
        fig.add_artist(Rectangle((lx, 0.849), 0.011, 0.017, transform=fig.transFigure, facecolor=cell_color(v),
                                 edgecolor="#d0d0d0" if v == 0 else cell_color(v)))
        t = fig.text(lx + 0.017, 0.857, lab, fontproperties=F["400"], fontsize=13, color=INK2, va="center")
        fig.canvas.draw(); lx = t.get_window_extent().transformed(fig.transFigure.inverted()).x1 + 0.03
    # layout
    left, right, top, bottom = 0.235, 0.955, 0.725, 0.165
    nrows = sum(len(d) + 1 for _, _, d in ordered)
    gap = 0.25
    total = nrows + gap * (len(ordered) - 1)
    rh = (top - bottom) / total
    cw = (right - left) / len(COLS)
    gw = 0.006  # gap between metric groups
    cw = (right - left - 2 * gw) / len(COLS)
    def cx(j):
        return left + j * cw + (j // 3) * gw
    # column headers
    for mi, m in enumerate(METRICS):
        x0 = cx(mi * 3); x1 = cx(mi * 3 + 2) + cw
        fig.text((x0 + x1) / 2, top + 0.033, MLABEL[m], ha="center", va="center", fontproperties=F["700"],
                 fontsize=12, color=INK)
        fig.add_artist(plt.Line2D([x0 + 0.004, x1 - 0.004], [top + 0.019, top + 0.019], color=RULE, linewidth=1,
                                  transform=fig.transFigure))
        for k, (_, w) in enumerate(WINDOWS):
            fig.text(cx(mi * 3 + k) + cw / 2, top + 0.006, w, ha="center", va="center", fontproperties=F["400"],
                     fontsize=11.5, color=INK2)
    y = top - 0.004
    for _, s, d in ordered:
        # sector median row
        y -= rh
        fig.text(0.05, y + rh / 2, SEC_NAME[s].upper() + "  ·  MEDIAN", fontproperties=F["700"], fontsize=11.5,
                 color=SEC_COLOR[s], va="center")
        for j, c in enumerate(COLS):
            v = d[c].median()
            fig.add_artist(Rectangle((cx(j) + 0.0015, y + 0.002), cw - 0.003, rh - 0.004, transform=fig.transFigure,
                                     facecolor=cell_color(v), edgecolor=SEC_COLOR[s], linewidth=1.4))
            fig.text(cx(j) + cw / 2, y + rh / 2, fmt(v) if not np.isnan(v) else "–", ha="center", va="center",
                     fontproperties=F["700"], fontsize=10.5, color=text_color(v))
        for _, r in d.iterrows():
            y -= rh
            fig.text(left - 0.01, y + rh / 2, r.Company, ha="right", va="center", fontproperties=F["400"],
                     fontsize=11.5, color=INK)
            for j, c in enumerate(COLS):
                v = r[c]
                fig.add_artist(Rectangle((cx(j) + 0.0015, y + 0.002), cw - 0.003, rh - 0.004,
                                         transform=fig.transFigure, facecolor=cell_color(v), edgecolor="none"))
                txt = fmt(v) if not np.isnan(v) else (r[c + "_flag"] or "")
                fig.text(cx(j) + cw / 2, y + rh / 2, txt, ha="center", va="center", fontproperties=F["400"],
                         fontsize=10.5, color=text_color(v) if not np.isnan(v) else MUTED)
        y -= rh * gap
    p = os.path.join(OUTDIR, "R1_consensus_revision_heatmap.png")
    fig.savefig(p, dpi=DPI, facecolor="white"); plt.close(fig)
    return p

# ---------------------------------------------------------------- 2) sector revision paths
def paths():
    dates = [B.SNAP + pd.Timedelta(days=lb) for lb in LOOKBACKS]
    res, excl = {}, {}
    for m in METRICS:
        for s in SEC_ORDER:
            series = {}
            for t in B.SECTORS[s]:
                if not B.has(t, m):
                    continue
                fy = B.fwd_year(t)
                v = np.array([cons(t, m, fy, lb) for lb in LOOKBACKS])
                if np.isnan(v[0]) or v[0] <= 0 or np.isnan(v).any():
                    excl.setdefault(m, []).append(B.SHORT[t]); continue
                series[t] = v / v[0] * 100
            res[(m, s)] = (pd.DataFrame(series, index=dates).median(axis=1), len(series))
    best = max(SEC_ORDER, key=lambda s: res[("op", s)][0].iloc[-1])
    title = (f"{SEC_NAME[best]} has had the biggest upgrades: median forward operating EPS consensus "
             f"{fmt(res[('op', best)][0].iloc[-1] / 100 - 1)} in a year")
    legend = [("line", SEC_NAME[s], (SEC_COLOR[s], "-")) for s in SEC_ORDER]
    ex_txt = "; ".join(f"{ {'rev': 'revenue', 'gaap': 'GAAP EPS', 'op': 'operating EPS'}[m]}: {', '.join(v)}"
                       for m, v in excl.items())
    fig = fig_base(title, legend, None, f"Source: Visible Alpha ({B.SNAP:%-d %b %Y}).",
                   "Note: Each company's forward fiscal-year consensus indexed to 100 one year earlier; line = median "
                   "of the group. Forward FY = FY28 for January year-ends, FY27 for the rest. Excluded for a negative or "
                   f"missing starting value — {ex_txt}. Consensus points at 365, 180, 120, 90, 60, 30, 14 and 7 days "
                   "before the snapshot.")
    titles = {"rev": "REVENUE CONSENSUS", "gaap": "GAAP EPS CONSENSUS", "op": "OPERATING EPS CONSENSUS"}
    ymax = max(res[k][0].max() for k in res) * 1.1
    ymin = min(90, min(res[k][0].min() for k in res) * 0.95)
    for i, m in enumerate(METRICS):
        ax = fig.add_axes([0.07 + i * 0.31, 0.21, 0.24, 0.53])
        fig.text(0.05 + i * 0.31, 0.81, titles[m] + ", NOV 25 = 100", fontproperties=F["700"], fontsize=11.5,
                 color=INK, va="center")
        ax.axhline(100, color="#9a9a9a", linewidth=1, linestyle=(0, (4, 3)), zorder=1)
        fig.canvas.draw(); tr = ax.transData; ends = []
        for s in SEC_ORDER:
            ser, k = res[(m, s)]
            if k == 0: continue
            ax.plot(ser.index, ser.values, color=SEC_COLOR[s], linewidth=3, marker="o", markersize=4.5, zorder=4,
                    linestyle="-" if k >= 2 else (0, (3, 2)))
            ends.append((ser.iloc[-1], s, k))
        ax.set_ylim(ymin, ymax)
        ax.set_xlim(dates[0] - pd.Timedelta(days=12), dates[-1] + pd.Timedelta(days=12))
        fig.canvas.draw(); tr = ax.transData; placed = []
        for v, s, k in sorted(ends):
            yp = tr.transform((mdates.date2num(dates[-1]), v))[1]
            if placed and yp - placed[-1] < 24: yp = placed[-1] + 24
            placed.append(yp)
            ax.annotate(f"{v:,.0f}" + (" (1 co.)" if k < 2 else ""), (dates[-1], v),
                        xytext=(dates[-1] + pd.Timedelta(days=22), tr.inverted().transform((0, yp))[1]),
                        textcoords="data", va="center", fontproperties=F["700"], fontsize=12.5, color=SEC_COLOR[s],
                        annotation_clip=False)
        tk = [dates[0], B.SNAP - pd.DateOffset(months=8), B.SNAP - pd.DateOffset(months=4), dates[-1]]
        ax.set_xticks(tk, [d.strftime("%b %y") for d in tk])
        ax.yaxis.set_major_locator(plt.MaxNLocator(6))
        ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f"{v:,.0f}"))
        style_axes(ax); relabel(ax)
    p = os.path.join(OUTDIR, "R2_sector_consensus_revision_paths.png")
    fig.savefig(p, dpi=DPI, facecolor="white"); plt.close(fig)
    return p, res, excl

if __name__ == "__main__":
    print(heatmap())
    p, res, excl = paths(); print(p)
    for (m, s), (ser, k) in res.items():
        print(m, s, k, ser.round(0).tolist())
    print(excl)
    out = R[["Sector", "Ticker", "Company", "Forward_FY"] + COLS].copy()
    out.to_excel(os.path.join(OUTDIR, "consensus_revisions.xlsx"), index=False)
    print(out.round(3).to_string())
