"""Write the deliverables: tables (CSV), machine-readable results (JSON) and
figures (PNG). Figure styling follows a single validated palette."""
from __future__ import annotations

import json
from dataclasses import dataclass
from datetime import date
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
import numpy as np  # noqa: E402
import pandas as pd  # noqa: E402

from .curves import svensson_forward, svensson_zero  # noqa: E402

# Palette (validated categorical slots + chrome) shared with the web page.
BLUE = "#2a78d6"
ORANGE = "#eb6834"
AQUA = "#1baf7a"
MUTED = "#898781"
GRID = "#e1e0d9"
AXIS = "#c3c2b7"
INK = "#0b0b0b"
INK_SECONDARY = "#52514e"
SURFACE = "#fcfcfb"
FIGSIZE = (9.5, 5.2)
DPI = 160
CURVE_GRID_POINTS = 600


def _style_axes(ax) -> None:
    ax.set_facecolor(SURFACE)
    for side in ("top", "right"):
        ax.spines[side].set_visible(False)
    for side in ("left", "bottom"):
        ax.spines[side].set_color(AXIS)
    ax.grid(True, axis="y", color=GRID, linewidth=0.8)
    ax.grid(False, axis="x")
    ax.tick_params(colors=INK_SECONDARY, labelsize=9)
    ax.xaxis.label.set_color(INK_SECONDARY)
    ax.yaxis.label.set_color(INK_SECONDARY)
    ax.title.set_color(INK)


def plot_zero_curve(bond_table: pd.DataFrame, params: np.ndarray, quote_date: date, settlement: date, path: Path) -> None:
    """Fitted zero and instantaneous forward curves over the bonds' street yields."""
    t_max = float(bond_table["years_to_maturity"].max()) + 0.5
    grid = np.linspace(0.02, t_max, CURVE_GRID_POINTS)
    zero = svensson_zero(grid, *params) * 100.0
    fwd = svensson_forward(grid, *params) * 100.0

    fig, ax = plt.subplots(figsize=FIGSIZE, facecolor=SURFACE)
    _style_axes(ax)
    used = bond_table[bond_table["in_fit"]]
    dropped = bond_table[~bond_table["in_fit"]]
    ax.scatter(used["years_to_maturity"], used["street_yield_pct"], s=14, color=MUTED, alpha=0.75,
               linewidths=0, label=f"Treasury street yields, in fit ({len(used)})", zorder=2)
    if len(dropped):
        ax.scatter(dropped["years_to_maturity"], dropped["street_yield_pct"], s=22, facecolors="none",
                   edgecolors=MUTED, linewidths=1.0, label=f"Excluded < 3 months ({len(dropped)})", zorder=2)
    ax.plot(grid, zero, color=BLUE, linewidth=2.0, label="Zero-coupon rate r(t), continuous", zorder=3)
    ax.plot(grid, fwd, color=ORANGE, linewidth=2.0, linestyle=(0, (5, 3)), label="Instantaneous forward f(t)", zorder=3)
    r30 = float(svensson_zero(30.0, *params) * 100.0)
    ax.annotate(f"r(30y) = {r30:.2f}%", xy=(30.0, r30), xytext=(8, -2), textcoords="offset points",
                va="center", fontsize=9, color=INK_SECONDARY)
    ax.set_xlabel("Years from settlement")
    ax.set_ylabel("Rate (% per year)")
    ax.set_xlim(0, t_max + 2.5)
    ax.set_title(f"US Treasury ZCB term structure  |  WSJ quotes {quote_date:%b %d, %Y}, settle {settlement:%b %d, %Y}",
                 fontsize=11, loc="left")
    ax.legend(frameon=False, fontsize=9, loc="lower right")
    fig.tight_layout()
    fig.savefig(path, dpi=DPI, facecolor=SURFACE)
    plt.close(fig)


def plot_residuals(bond_table: pd.DataFrame, summary: dict, path: Path) -> None:
    """Yield fitting errors (model minus market) in basis points by maturity."""
    fig, ax = plt.subplots(figsize=FIGSIZE, facecolor=SURFACE)
    _style_axes(ax)
    used = bond_table[bond_table["in_fit"]]
    ax.axhline(0, color=AXIS, linewidth=1.0)
    rmse = summary["yield_rmse_bp"]
    ax.axhspan(-rmse, rmse, color=BLUE, alpha=0.08, linewidth=0, label=f"+/- RMSE ({rmse:.1f} bp)")
    ax.scatter(used["years_to_maturity"], used["yield_error_bp"], s=16, color=BLUE, alpha=0.8, linewidths=0,
               label="Model yield - market yield (bp)")
    ax.set_xlabel("Years to maturity")
    ax.set_ylabel("Yield error (bp)")
    ax.set_title("Fitting errors of the Svensson curve, bonds used in the fit", fontsize=11, loc="left")
    ax.legend(frameon=False, fontsize=9, loc="upper right")
    fig.tight_layout()
    fig.savefig(path, dpi=DPI, facecolor=SURFACE)
    plt.close(fig)


def plot_discount_factors(payment_table: pd.DataFrame, path: Path) -> None:
    """Discount factor at every Treasury payment date."""
    fig, ax = plt.subplots(figsize=FIGSIZE, facecolor=SURFACE)
    _style_axes(ax)
    ax.plot(payment_table["t_years"], payment_table["discount_factor"], color=AQUA, linewidth=2.0)
    ax.scatter(payment_table["t_years"], payment_table["discount_factor"], s=6, color=AQUA, linewidths=0)
    last = payment_table.iloc[-1]
    ax.annotate(f"d({last['t_years']:.1f}y) = {last['discount_factor']:.4f}", xy=(last["t_years"], last["discount_factor"]),
                xytext=(6, 6), textcoords="offset points", fontsize=9, color=INK_SECONDARY)
    ax.set_xlabel("Years from settlement")
    ax.set_ylabel("Discount factor d(t) = exp(-r(t) t)")
    ax.set_ylim(0, 1.02)
    ax.set_title(f"Continuous discount factors at all {len(payment_table)} Treasury payment dates", fontsize=11, loc="left")
    fig.tight_layout()
    fig.savefig(path, dpi=DPI, facecolor=SURFACE)
    plt.close(fig)


def _json_ready(value):
    if isinstance(value, (np.floating, np.integer)):
        return value.item()
    if isinstance(value, (date, pd.Timestamp)):
        return value.isoformat()
    if isinstance(value, np.ndarray):
        return value.tolist()
    if isinstance(value, dict):
        return {k: _json_ready(v) for k, v in value.items()}
    if isinstance(value, (list, tuple)):
        return [_json_ready(v) for v in value]
    return value


def frame_records(frame: pd.DataFrame) -> list[dict]:
    return [_json_ready(r) for r in frame.to_dict(orient="records")]


def write_json(payload: dict, path: Path) -> None:
    path.write_text(json.dumps(_json_ready(payload), indent=2))
