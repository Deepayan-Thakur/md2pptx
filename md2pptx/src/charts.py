"""
charts.py — Programmatic chart generation using matplotlib.
All charts use the Accenture theme palette. Returns PNG bytes.
"""
import io
import re
import math
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np
from typing import List, Dict, Any, Optional, Tuple

# Accenture brand palette
COLORS = ["#EF4444", "#0F9ED5", "#196B24", "#E97132", "#A02B93", "#2C2C2C", "#4EA72E"]
BG_MAIN = "#FFFFFF"
BG_PLOT = "#F7F7F7"
TEXT_DARK = "#2C2C2C"
GRID_COLOR = "#E5E5E5"
FONT_TITLE = "DejaVu Sans"
FONT_BODY = "DejaVu Sans"


def _style_axes(ax, title: str):
    ax.set_facecolor(BG_PLOT)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_color(GRID_COLOR)
    ax.spines["bottom"].set_color(GRID_COLOR)
    ax.tick_params(colors=TEXT_DARK, labelsize=10)
    ax.yaxis.grid(True, color=GRID_COLOR, linewidth=0.8, zorder=0)
    ax.set_axisbelow(True)
    ax.set_title(title, fontsize=14, fontweight="bold", color=TEXT_DARK,
                 fontfamily=FONT_TITLE, pad=12)


def _to_bytes(fig) -> bytes:
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight",
                facecolor=BG_MAIN, edgecolor="none")
    plt.close(fig)
    buf.seek(0)
    return buf.read()


def _parse_values(values) -> List[float]:
    result = []
    for v in values:
        try:
            clean = re.sub(r"[,$%\s~≈]", "", str(v))
            result.append(float(clean))
        except (ValueError, TypeError):
            result.append(0.0)
    return result


def bar_chart(labels: List[str], values, ylabel: str = "",
              chart_title: str = "", horizontal: bool = False) -> bytes:
    values = _parse_values(values)
    if not values:
        return b""

    n = len(values)
    fig_w = max(8, min(12, n * 1.4))
    fig, ax = plt.subplots(figsize=(fig_w, 5.5))
    fig.patch.set_facecolor(BG_MAIN)
    x = np.arange(n)

    if horizontal:
        bars = ax.barh(x, values, color=[COLORS[i % len(COLORS)] for i in range(n)],
                       height=0.6, edgecolor="white", linewidth=1.5, zorder=3)
        ax.set_yticks(x)
        ax.set_yticklabels(labels, fontsize=10, color=TEXT_DARK)
        ax.xaxis.grid(True, color=GRID_COLOR, linewidth=0.8, zorder=0)
        ax.yaxis.grid(False)
        for bar, val in zip(bars, values):
            ax.text(bar.get_width() + max(values) * 0.01,
                    bar.get_y() + bar.get_height() / 2,
                    f"{val:,.1f}", va="center", fontsize=9,
                    fontweight="bold", color=TEXT_DARK)
    else:
        bars = ax.bar(x, values, color=[COLORS[i % len(COLORS)] for i in range(n)],
                      width=0.6, edgecolor="white", linewidth=1.5, zorder=3)
        ax.set_xticks(x)
        ax.set_xticklabels(labels, fontsize=10, color=TEXT_DARK,
                           rotation=20 if n > 5 else 0, ha="right" if n > 5 else "center")
        if ylabel:
            ax.set_ylabel(ylabel, fontsize=11, color="#666666")
        for bar, val in zip(bars, values):
            ax.text(bar.get_x() + bar.get_width() / 2,
                    bar.get_height() + max(values) * 0.012,
                    f"{val:,.1f}", ha="center", fontsize=9,
                    fontweight="bold", color=TEXT_DARK)

    _style_axes(ax, chart_title)
    plt.tight_layout(pad=1.2)
    return _to_bytes(fig)


def line_chart(labels: List[str], values, ylabel: str = "",
               chart_title: str = "") -> bytes:
    values = _parse_values(values)
    if not values:
        return b""
    fig, ax = plt.subplots(figsize=(10, 5))
    fig.patch.set_facecolor(BG_MAIN)
    x = np.arange(len(labels))
    ax.plot(x, values, color=COLORS[0], linewidth=2.5, marker="o",
            markersize=8, markerfacecolor="white", markeredgewidth=2.5,
            markeredgecolor=COLORS[0], zorder=4)
    ax.fill_between(x, values, alpha=0.08, color=COLORS[0])
    ax.set_xticks(x)
    ax.set_xticklabels(labels, fontsize=10, color=TEXT_DARK)
    if ylabel:
        ax.set_ylabel(ylabel, fontsize=11, color="#666666")
    for xi, val in zip(x, values):
        ax.annotate(f"{val:,.1f}", (xi, val), textcoords="offset points",
                    xytext=(0, 10), ha="center", fontsize=9,
                    fontweight="bold", color=TEXT_DARK)
    _style_axes(ax, chart_title)
    plt.tight_layout(pad=1.2)
    return _to_bytes(fig)


def pie_chart(labels: List[str], values, chart_title: str = "") -> bytes:
    values = _parse_values(values)
    if not values or sum(values) == 0:
        return b""
    n = len(values)
    col = COLORS[:n] if n <= len(COLORS) else [COLORS[i % len(COLORS)] for i in range(n)]
    fig, ax = plt.subplots(figsize=(9, 5.5))
    fig.patch.set_facecolor(BG_MAIN)
    wedges, texts, autotexts = ax.pie(
        values, labels=None, autopct="%1.1f%%",
        colors=col, startangle=90, pctdistance=0.78,
        wedgeprops={"linewidth": 2, "edgecolor": "white"},
    )
    for at in autotexts:
        at.set_fontsize(10)
        at.set_fontweight("bold")
        at.set_color("white")
    # Donut hole
    ax.add_artist(plt.Circle((0, 0), 0.48, fc="white"))
    ax.legend(wedges, labels, loc="center left", bbox_to_anchor=(1, 0, 0.5, 1),
              fontsize=10, frameon=False)
    ax.set_title(chart_title, fontsize=14, fontweight="bold", color=TEXT_DARK,
                 fontfamily=FONT_TITLE, pad=10)
    plt.tight_layout(pad=1.2)
    return _to_bytes(fig)


def grouped_bar_chart(labels: List[str], series: Dict[str, List],
                      chart_title: str = "") -> bytes:
    """Multi-series bar chart."""
    n = len(labels)
    m = len(series)
    fig, ax = plt.subplots(figsize=(max(9, n * m * 0.5 + 2), 5.5))
    fig.patch.set_facecolor(BG_MAIN)
    x = np.arange(n)
    width = 0.7 / m
    for i, (name, vals) in enumerate(series.items()):
        vals = _parse_values(vals)
        offset = (i - m / 2 + 0.5) * width
        ax.bar(x + offset, vals, width * 0.9,
               label=name, color=COLORS[i % len(COLORS)],
               edgecolor="white", linewidth=1.2, zorder=3)
    ax.set_xticks(x)
    ax.set_xticklabels(labels, fontsize=10, color=TEXT_DARK)
    ax.legend(fontsize=10, frameon=False)
    _style_axes(ax, chart_title)
    plt.tight_layout(pad=1.2)
    return _to_bytes(fig)


def stat_card_image(stats: List[Tuple[str, str, str]],
                    title: str = "") -> bytes:
    """
    stats = [(value, label, delta), ...]   e.g. ("$5.9B", "AI Bookings", "+100%")
    Returns a PNG with styled stat cards side by side.
    """
    n = min(len(stats), 4)
    if n == 0:
        return b""
    fig_w = n * 2.8 + 0.6
    fig, axes = plt.subplots(1, n, figsize=(fig_w, 2.8))
    fig.patch.set_facecolor(BG_MAIN)
    if n == 1:
        axes = [axes]

    for i, (ax, (value, label, delta)) in enumerate(zip(axes, stats[:n])):
        ax.set_facecolor(COLORS[i % len(COLORS)])
        ax.set_xlim(0, 1)
        ax.set_ylim(0, 1)
        ax.axis("off")
        ax.text(0.5, 0.65, value, ha="center", va="center",
                fontsize=22, fontweight="bold", color="white", transform=ax.transAxes)
        ax.text(0.5, 0.32, label, ha="center", va="center",
                fontsize=10, color="white", transform=ax.transAxes, wrap=True)
        if delta:
            ax.text(0.5, 0.12, delta, ha="center", va="center",
                    fontsize=9, color="white", alpha=0.85, transform=ax.transAxes)

    if title:
        fig.suptitle(title, fontsize=13, fontweight="bold", color=TEXT_DARK, y=1.02)
    plt.tight_layout(pad=0.3)
    return _to_bytes(fig)


def area_chart(labels: List[str], values, ylabel: str = "",
               chart_title: str = "") -> bytes:
    """Filled line chart."""
    values = _parse_values(values)
    if not values:
        return b""
    fig, ax = plt.subplots(figsize=(10, 5.2))
    fig.patch.set_facecolor(BG_MAIN)
    x = np.arange(len(labels))
    ax.plot(x, values, color=COLORS[1], linewidth=3, zorder=4)
    ax.fill_between(x, values, alpha=0.25, color=COLORS[1], zorder=3)
    ax.set_xticks(x)
    ax.set_xticklabels(labels, fontsize=10, color=TEXT_DARK)
    if ylabel:
        ax.set_ylabel(ylabel, fontsize=11, color="#666666")
    _style_axes(ax, chart_title)
    plt.tight_layout(pad=1.2)
    return _to_bytes(fig)


def render_chart(chart_spec: Dict) -> bytes:
    """Dispatch to the right chart type from a spec dict."""
    if not chart_spec:
        return b""
    ctype = chart_spec.get("type", "bar").lower()
    labels = chart_spec.get("labels", [])
    values = chart_spec.get("values", [])
    title = chart_spec.get("chart_title", chart_spec.get("title", ""))
    ylabel = chart_spec.get("ylabel", "")

    if not labels or not values:
        return b""
    
    # Auto-dispatch logic based on data characteristics
    if ctype == "auto" or not ctype:
        avg = sum(_parse_values(values)) / len(values) if values else 0
        if any(keyword in title.lower() for keyword in ["trend", "growth", "over time", "evolution"]):
            ctype = "line"
        elif any(keyword in title.lower() for keyword in ["share", "distribution", "breakdown", "portfolio"]):
            ctype = "pie"
        else:
            ctype = "bar"

    if ctype == "pie":
        return pie_chart(labels, values, title)
    if ctype == "line":
        return line_chart(labels, values, ylabel, title)
    if ctype == "area":
        return area_chart(labels, values, ylabel, title)
    
    return bar_chart(labels, values, ylabel, title,
                     horizontal=len(labels) > 6)