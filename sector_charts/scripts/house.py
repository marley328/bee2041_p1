"""House chart style (matches the Kolytics examples)."""
import os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties, fontManager

FONT_DIR = os.path.expanduser("~/.fonts")
for w in ("400", "600", "700"):
    fontManager.addfont(os.path.join(FONT_DIR, f"SourceSans3-{w}.ttf"))
F = {w: FontProperties(fname=os.path.join(FONT_DIR, f"SourceSans3-{w}.ttf")) for w in ("400", "600", "700")}

TITLE = "#2f6b71"      # headline teal
INK = "#2b2b33"        # primary text
INK2 = "#555560"       # axis labels, secondary
MUTED = "#8a8a8a"      # footer
GRID = "#e6e6e6"
RULE = "#cfcfcf"
REPORTED = "#808080"
CONSENSUS = "#bcbab5"
EDGE = "#3a3a3a"
TEAL_DARK = "#336e74"
TEAL_LIGHT = "#82a9aa"
SALMON = "#f68a7a"
BLACK = "#26262e"
SHADE = "#f2f2f2"

W, H, DPI = 12, 8, 167  # ~2000 x 1336 px, as in the examples


def fig_base(title, legend=None, subtitle=None, source="", note=""):
    fig = plt.figure(figsize=(W, H), dpi=DPI)
    fig.patch.set_facecolor("white")
    tt = fig.text(0.05, 0.94, title, fontproperties=F["700"], fontsize=24, color=TITLE, va="center")
    fig.canvas.draw()
    fs = 24
    while fs > 16 and tt.get_window_extent().transformed(fig.transFigure.inverted()).x1 > 0.955:
        fs -= 1; tt.set_fontsize(fs); fig.canvas.draw()
    if legend:
        x = 0.05
        for kind, label, style in legend:
            x = _legend_item(fig, x, 0.857, kind, label, style)
    if subtitle:
        fig.text(0.05, 0.81, subtitle.upper(), fontproperties=F["700"], fontsize=11.5, color=INK, va="center")
    fig.add_artist(plt.Line2D([0.05, 0.955], [0.1, 0.1], color=RULE, linewidth=1, transform=fig.transFigure))
    import textwrap
    lines = [source] if source else []
    if note:
        lines += textwrap.wrap(note, 178)
    y = 0.078
    for line in lines:
        fig.text(0.05, y, line, fontproperties=F["400"], fontsize=10, color=MUTED, va="center")
        y -= 0.019
    return fig


def _legend_item(fig, x, y, kind, label, style):
    from matplotlib.patches import Rectangle
    if kind == "box":
        fc, ec, ls = style
        fig.add_artist(Rectangle((x, y - 0.008), 0.011, 0.017, transform=fig.transFigure, facecolor=fc,
                                 edgecolor=ec or fc, linestyle=ls or "-", linewidth=1.1))
        tx = x + 0.017
    else:  # line
        color, ls = style
        fig.add_artist(plt.Line2D([x, x + 0.026], [y, y], color=color, linewidth=2.6, linestyle=ls,
                                  transform=fig.transFigure))
        tx = x + 0.032
    t = fig.text(tx, y, label, fontproperties=F["400"], fontsize=13, color=INK2, va="center")
    fig.canvas.draw()
    bb = t.get_window_extent().transformed(fig.transFigure.inverted())
    return bb.x1 + 0.03


def style_axes(ax, ygrid=True, xgrid=False):
    for s in ax.spines.values():
        s.set_visible(False)
    ax.tick_params(length=0, labelsize=13, colors=INK2, pad=8)
    for lab in ax.get_xticklabels() + ax.get_yticklabels():
        lab.set_fontproperties(F["400"]); lab.set_fontsize(13)
    if ygrid:
        ax.grid(axis="y", color=GRID, linewidth=1)
    if xgrid:
        ax.grid(axis="x", color=GRID, linewidth=1)
    ax.set_axisbelow(True)


def relabel(ax):
    for lab in ax.get_xticklabels() + ax.get_yticklabels():
        lab.set_fontproperties(F["400"]); lab.set_fontsize(13); lab.set_color(INK2)
