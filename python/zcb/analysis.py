"""Glue between the quote sheet, bond mechanics and the curve fit."""
from __future__ import annotations

from dataclasses import dataclass
from datetime import date, timedelta

import numpy as np
import pandas as pd

from .bonds import accrued_interest, build_schedule, modified_duration, street_yield
from .curves import (
    BondPricingInput,
    FitResult,
    discount_factor,
    fit_curve,
    svensson_forward,
    svensson_zero,
)
from .quotes import parse_wsj_price, parse_wsj_price_decimal_tenths

BASIS_POINTS = 1e4
PRICE_COLUMNS = {"ask": "ask_price", "bid": "bid_price", "mid": "mid_price"}
# Bonds inside the money-market segment are excluded from the fit (Gurkaynak,
# Sack & Wright 2007 use the same 3-month cut-off): their prices reflect
# funding conditions, and with near-zero duration a 1-cent price error is a
# 50 bp yield error.
DEFAULT_MIN_YEARS_IN_FIT = 0.25

# SIFMA bond-market holidays for the rest of 2026 (the data is a Sep 2026 quote sheet).
# Extend this set before using the settlement helper on quote sheets from other years.
US_MARKET_HOLIDAYS = {
    date(2026, 9, 7),   # Labor Day 2026
    date(2026, 10, 12),  # Columbus Day 2026
    date(2026, 11, 11),  # Veterans Day 2026
    date(2026, 11, 26),  # Thanksgiving 2026
    date(2026, 12, 25),  # Christmas 2026
}


def settlement_date_for(quote_date: date, business_days: int = 1) -> date:
    """Regular-way Treasury settlement: T+1 business days, skipping weekends and holidays."""
    current = quote_date
    remaining = business_days
    while remaining > 0:
        current += timedelta(days=1)
        if current.weekday() < 5 and current not in US_MARKET_HOLIDAYS:
            remaining -= 1
    return current


@dataclass(frozen=True)
class PreparedBonds:
    settlement: date
    price_side: str
    min_years_in_fit: float
    table: pd.DataFrame  # every bond, with an ``in_fit`` flag
    pricing_inputs: list[BondPricingInput]  # only the bonds used in the fit


def _bond_label(maturity: date, coupon: float) -> str:
    return f"{coupon:.3f}% {maturity.isoformat()}"


def prepare_bonds(
    quotes: pd.DataFrame,
    settlement: date,
    price_side: str = "ask",
    min_years_in_fit: float = DEFAULT_MIN_YEARS_IN_FIT,
) -> PreparedBonds:
    """Per-bond mechanics: schedule, accrued, dirty price, street yield, duration weight."""
    if price_side not in PRICE_COLUMNS:
        raise ValueError(f"price_side must be one of {sorted(PRICE_COLUMNS)}")
    price_col = PRICE_COLUMNS[price_side]
    rows, inputs = [], []
    for q in quotes.itertuples(index=False):
        schedule = build_schedule(q.maturity, q.coupon, settlement)
        accrued = accrued_interest(q.coupon, settlement, schedule)
        clean = float(getattr(q, price_col))
        dirty = clean + accrued
        ytm = street_yield(dirty, settlement, schedule)
        duration = modified_duration(ytm, settlement, schedule)
        label = _bond_label(q.maturity, q.coupon)
        rows.append(
            {
                "label": label,
                "maturity": q.maturity,
                "coupon": q.coupon,
                "bid_quote": q.bid_quote,
                "ask_quote": q.ask_quote,
                "clean_price": clean,
                "wsj_asked_yield": q.asked_yield,
                "previous_coupon": schedule.previous_coupon,
                "next_coupon": schedule.next_coupon,
                "n_cash_flows": len(schedule.amounts),
                "accrued_interest": accrued,
                "dirty_price": dirty,
                "street_yield_pct": ytm * 100.0,
                "modified_duration": duration,
                "years_to_maturity": schedule.times[-1],
                "in_fit": schedule.times[-1] >= min_years_in_fit,
            }
        )
        if schedule.times[-1] < min_years_in_fit:
            continue
        inputs.append(
            BondPricingInput(
                label=label,
                times=np.asarray(schedule.times),
                amounts=np.asarray(schedule.amounts),
                market_dirty_price=dirty,
                # (P_model - P_mkt) / (P_mkt x D_mod) is the first-order yield error, so the
                # least-squares objective is the sum of squared yield errors (slide 31).
                weight=1.0 / (dirty * duration),
            )
        )
    table = pd.DataFrame(rows)
    return PreparedBonds(settlement, price_side, min_years_in_fit, table, inputs)


def check_conventions(quotes: pd.DataFrame, quote_date: date) -> pd.DataFrame:
    """RMSE (bp) of our recomputed asked yield vs the WSJ 'Asked yield' column
    under alternative settlement dates and price-format readings.

    Used to justify the T+1 settlement date and the 32nds-with-eighths reading.
    """
    candidates = {
        "T+0 (quote date)": quote_date,
        "T+1 calendar": quote_date + timedelta(days=1),
        "T+1 business (regular way)": settlement_date_for(quote_date, 1),
        "T+2 business": settlement_date_for(quote_date, 2),
    }
    parsers = {
        "32nds + eighths": parse_wsj_price,
        "32nds + tenths": parse_wsj_price_decimal_tenths,
        "plain decimal": float,
    }
    records = []
    for settle_label, settlement in candidates.items():
        for parse_label, parser in parsers.items():
            errors = []
            for q in quotes.itertuples(index=False):
                schedule = build_schedule(q.maturity, q.coupon, settlement)
                dirty = parser(q.ask_quote) + accrued_interest(q.coupon, settlement, schedule)
                ytm = street_yield(dirty, settlement, schedule) * 100.0
                errors.append(ytm - q.asked_yield)
            errors = np.asarray(errors) * 100.0  # percentage points -> bp
            records.append(
                {
                    "settlement_rule": settle_label,
                    "settlement": settlement,
                    "price_format": parse_label,
                    "rmse_bp": float(np.sqrt(np.mean(errors**2))),
                    "median_abs_bp": float(np.median(np.abs(errors))),
                    "max_abs_bp": float(np.max(np.abs(errors))),
                }
            )
    return pd.DataFrame(records).sort_values("rmse_bp").reset_index(drop=True)


@dataclass
class TermStructureFit:
    fit: FitResult
    bonds: PreparedBonds
    bond_table: pd.DataFrame  # bonds.table + model columns
    summary: dict


def fit_term_structure(bonds: PreparedBonds, model: str = "svensson") -> TermStructureFit:
    """Fit the curve and evaluate price and yield errors bond by bond."""
    fit = fit_curve(bonds.pricing_inputs, model=model)
    table = bonds.table.copy()
    model_dirty, model_yields = [], []
    for q in table.itertuples(index=False):
        schedule = build_schedule(q.maturity, q.coupon, bonds.settlement)
        dirty = float(np.dot(schedule.amounts, discount_factor(np.asarray(schedule.times), fit.params)))
        model_dirty.append(dirty)
        model_yields.append(street_yield(dirty, bonds.settlement, schedule) * 100.0)
    table["model_dirty_price"] = model_dirty
    table["model_clean_price"] = table["model_dirty_price"] - table["accrued_interest"]
    table["price_error"] = table["model_clean_price"] - table["clean_price"]
    table["model_yield_pct"] = model_yields
    table["yield_error_bp"] = (table["model_yield_pct"] - table["street_yield_pct"]) * 100.0

    fitted = table[table["in_fit"]]
    yerr = fitted["yield_error_bp"].to_numpy()
    perr = fitted["price_error"].to_numpy()
    summary = {
        "model": model,
        "n_bonds_total": int(len(table)),
        "n_bonds_in_fit": int(len(fitted)),
        "min_years_in_fit": bonds.min_years_in_fit,
        "settlement": bonds.settlement.isoformat(),
        "price_side": bonds.price_side,
        "params": fit.params_dict,
        "weighted_sse": fit.weighted_sse,
        "price_rmse": float(np.sqrt(np.mean(perr**2))),
        "price_mae": float(np.mean(np.abs(perr))),
        "yield_rmse_bp": float(np.sqrt(np.mean(yerr**2))),
        "yield_mae_bp": float(np.mean(np.abs(yerr))),
        "yield_max_abs_bp": float(np.max(np.abs(yerr))),
        "n_starts": fit.n_starts,
        "n_starts_at_optimum": int(sum(r["sse"] <= fit.weighted_sse * (1 + 1e-6) for r in fit.start_results)),
        "converged": bool(fit.converged),
        "bound_active_beta0": bool(fit.params[0] <= 1e-9),
    }
    return TermStructureFit(fit, bonds, table, summary)


def payment_date_table(bonds: PreparedBonds, params: np.ndarray) -> pd.DataFrame:
    """The deliverable: continuously compounded discount rate for every
    Treasury payment date (coupon and principal) in the sample."""
    all_dates: set[date] = set()
    for q in bonds.table.itertuples(index=False):
        schedule = build_schedule(q.maturity, q.coupon, bonds.settlement)
        all_dates.update(schedule.dates)
    dates = sorted(all_dates)
    t = np.array([(d - bonds.settlement).days / 365.0 for d in dates])
    zero = svensson_zero(t, *params)
    forward = svensson_forward(t, *params)
    return pd.DataFrame(
        {
            "payment_date": dates,
            "days": [(d - bonds.settlement).days for d in dates],
            "t_years": t,
            "zero_rate_cc_pct": zero * 100.0,
            "discount_factor": discount_factor(t, params),
            "inst_forward_cc_pct": forward * 100.0,
        }
    )


BETA0_PROFILE_GRID = (0.0, 0.01, 0.02, 0.03, 0.04)


def profile_beta0(bonds: PreparedBonds, grid=BETA0_PROFILE_GRID) -> pd.DataFrame:
    """Refit with beta0 pinned at each grid value: shows how weakly the
    asymptotic level is identified by a 30-year sample, and how little the
    in-sample curve moves."""
    horizons = (0.25, 1.0, 5.0, 10.0, 20.0, 30.0)
    records = []
    for beta0 in grid:
        fit = fit_curve(bonds.pricing_inputs, "svensson", fixed_params={"beta0": beta0})
        yerr = _approx_yield_errors_bp(bonds, fit.params)
        record = {
            "beta0_fixed_pct": beta0 * 100.0,
            "weighted_sse": fit.weighted_sse,
            "approx_yield_rmse_bp": float(np.sqrt(np.mean(yerr**2))),
            **{f"beta{i}": float(fit.params[i]) for i in (1, 2, 3)},
            "tau1": float(fit.params[4]),
            "tau2": float(fit.params[5]),
        }
        for h in horizons:
            record[f"zero_{h:g}y_pct"] = float(svensson_zero(h, *fit.params) * 100.0)
        records.append(record)
    return pd.DataFrame(records)


def _approx_yield_errors_bp(bonds: PreparedBonds, params: np.ndarray) -> np.ndarray:
    """First-order yield error: (model - market dirty price) / (dirty price x duration) = weighted residual."""
    errors = []
    for b in bonds.pricing_inputs:
        model_dirty = float(np.dot(b.amounts, discount_factor(b.times, params)))
        errors.append((model_dirty - b.market_dirty_price) * b.weight)
    return np.asarray(errors) * BASIS_POINTS
