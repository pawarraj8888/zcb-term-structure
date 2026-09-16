"""Parametric zero-coupon curves (Svensson 1994, Nelson-Siegel 1987) and the
least-squares fit to coupon-bond prices.

All rates are continuously compounded decimals (0.045 = 4.5%); times are years.
"""
from __future__ import annotations

import itertools
from dataclasses import dataclass, field

import numpy as np
from scipy.optimize import least_squares

PARAM_NAMES = ("beta0", "beta1", "beta2", "beta3", "tau1", "tau2")
TINY_TIME = 1e-12

# Optimiser bounds: betas are rate-like decimals, taus are years.
LOWER_BOUNDS = np.array([0.0, -1.0, -1.0, -1.0, 0.05, 0.05])
UPPER_BOUNDS = np.array([0.25, 1.0, 1.0, 1.0, 30.0, 30.0])
TAU1_STARTS = (0.5, 1.0, 2.0, 4.0)
TAU2_STARTS = (3.0, 6.0, 10.0, 15.0, 20.0)


def _loading(t: np.ndarray, tau: float) -> np.ndarray:
    """(1 - exp(-t/tau)) / (t/tau), with the t -> 0 limit of 1."""
    x = np.asarray(t, dtype=float) / tau
    safe_x = np.where(x < TINY_TIME, TINY_TIME, x)
    return np.where(x < TINY_TIME, 1.0, (1.0 - np.exp(-safe_x)) / safe_x)


def svensson_zero(t, beta0, beta1, beta2, beta3, tau1, tau2) -> np.ndarray:
    """Continuously compounded zero rate r(t) under the Svensson form."""
    t = np.asarray(t, dtype=float)
    load1 = _loading(t, tau1)
    load2 = _loading(t, tau2)
    return (
        beta0
        + beta1 * load1
        + beta2 * (load1 - np.exp(-t / tau1))
        + beta3 * (load2 - np.exp(-t / tau2))
    )


def svensson_forward(t, beta0, beta1, beta2, beta3, tau1, tau2) -> np.ndarray:
    """Instantaneous forward rate f(t) implied by the Svensson zero curve."""
    t = np.asarray(t, dtype=float)
    return (
        beta0
        + beta1 * np.exp(-t / tau1)
        + beta2 * (t / tau1) * np.exp(-t / tau1)
        + beta3 * (t / tau2) * np.exp(-t / tau2)
    )


def discount_factor(t, params: np.ndarray) -> np.ndarray:
    """d(t) = exp(-r(t) t)."""
    t = np.asarray(t, dtype=float)
    return np.exp(-svensson_zero(t, *params) * t)


@dataclass(frozen=True)
class BondPricingInput:
    """Everything the curve fit needs about one bond."""

    label: str
    times: np.ndarray
    amounts: np.ndarray
    market_dirty_price: float
    weight: float  # residual multiplier, e.g. 1 / modified duration


def model_dirty_price(bond: BondPricingInput, params: np.ndarray) -> float:
    return float(np.dot(bond.amounts, discount_factor(bond.times, params)))


@dataclass(frozen=True)
class _CashFlowMatrix:
    """Bonds stacked into padded (n_bonds x max_flows) arrays for fast pricing."""

    times: np.ndarray
    amounts: np.ndarray  # zero-padded, so padding never contributes
    prices: np.ndarray
    weights: np.ndarray

    @classmethod
    def from_bonds(cls, bonds: list[BondPricingInput]) -> "_CashFlowMatrix":
        width = max(len(b.times) for b in bonds)
        times = np.ones((len(bonds), width))  # padded times are 1.0 (any finite value)
        amounts = np.zeros((len(bonds), width))
        for i, b in enumerate(bonds):
            times[i, : len(b.times)] = b.times
            amounts[i, : len(b.amounts)] = b.amounts
        prices = np.array([b.market_dirty_price for b in bonds])
        weights = np.array([b.weight for b in bonds])
        return cls(times, amounts, prices, weights)

    def model_prices(self, params: np.ndarray) -> np.ndarray:
        return np.sum(self.amounts * discount_factor(self.times, params), axis=1)

    def weighted_residuals(self, params: np.ndarray) -> np.ndarray:
        return self.weights * (self.model_prices(params) - self.prices)


def weighted_residuals(params: np.ndarray, bonds: list[BondPricingInput]) -> np.ndarray:
    """Weighted price residuals for a parameter vector (convenience wrapper)."""
    return _CashFlowMatrix.from_bonds(bonds).weighted_residuals(np.asarray(params))


MODELS = {
    # model name -> Svensson parameters held fixed (name -> value)
    "svensson": {},
    "nelson_siegel": {"beta3": 0.0, "tau2": 1.0},
}


def _expand(free: np.ndarray, free_index: list[int], fixed: dict[str, float]) -> np.ndarray:
    """Map the optimiser's free parameters back to the full 6-vector."""
    full = np.empty(len(PARAM_NAMES))
    for name, value in fixed.items():
        full[PARAM_NAMES.index(name)] = value
    full[free_index] = free
    return full


@dataclass
class FitResult:
    model: str
    params: np.ndarray
    weighted_sse: float
    n_bonds: int
    n_starts: int
    converged: bool
    start_results: list[dict] = field(default_factory=list)

    @property
    def params_dict(self) -> dict[str, float]:
        return {name: float(v) for name, v in zip(PARAM_NAMES, self.params)}


def _initial_betas(bonds: list[BondPricingInput]) -> tuple[float, float]:
    """Rough level/slope guesses from the shortest and longest bonds' money yields."""
    by_maturity = sorted(bonds, key=lambda b: b.times[-1])
    short, long = by_maturity[0], by_maturity[-1]

    def crude_zero(b: BondPricingInput) -> float:
        total = float(b.amounts.sum())
        return float(np.log(total / b.market_dirty_price) / b.times[-1])

    level = crude_zero(long)
    return level, crude_zero(short) - level


def fit_curve(
    bonds: list[BondPricingInput],
    model: str = "svensson",
    fixed_params: dict[str, float] | None = None,
    tau1_starts=TAU1_STARTS,
    tau2_starts=TAU2_STARTS,
) -> FitResult:
    """Fit Svensson (6 params) or Nelson-Siegel (beta3 = 0) by weighted least squares.

    ``fixed_params`` pins additional parameters (e.g. ``{"beta0": 0.03}``) for
    profiling. Svensson objectives are multimodal in (tau1, tau2), so the fit is
    run from a grid of decay-parameter starting points and the best optimum kept.
    """
    if model not in MODELS:
        raise ValueError(f"Unknown model {model!r}; choose from {sorted(MODELS)}")
    fixed = {**MODELS[model], **(fixed_params or {})}
    unknown = set(fixed) - set(PARAM_NAMES)
    if unknown:
        raise ValueError(f"Unknown parameter names {sorted(unknown)}")
    free_index = [i for i, name in enumerate(PARAM_NAMES) if name not in fixed]
    matrix = _CashFlowMatrix.from_bonds(bonds)
    level, slope = _initial_betas(bonds)
    lower, upper = LOWER_BOUNDS[free_index], UPPER_BOUNDS[free_index]

    def residuals(free: np.ndarray) -> np.ndarray:
        return matrix.weighted_residuals(_expand(free, free_index, fixed))

    tau2_grid = (1.0,) if "tau2" in fixed else tau2_starts
    tau1_grid = (1.0,) if "tau1" in fixed else tau1_starts
    starts = []
    for tau1, tau2 in itertools.product(tau1_grid, tau2_grid):
        if "tau2" not in fixed and "tau1" not in fixed and tau2 <= tau1:
            continue
        full = np.array([level, slope, 0.0, 0.0, tau1, tau2])
        starts.append(np.clip(full[free_index], lower + 1e-9, upper - 1e-9))

    best = None
    start_results = []
    for x0 in starts:
        sol = least_squares(
            residuals,
            x0,
            bounds=(lower, upper),
            method="trf",
            x_scale="jac",
            xtol=1e-12,
            ftol=1e-12,
            gtol=1e-12,
            max_nfev=5000,
        )
        sse = float(np.sum(sol.fun**2))
        start_results.append({"start": x0.tolist(), "sse": sse, "status": int(sol.status)})
        if best is None or sse < best[0]:
            best = (sse, sol)
    sse, sol = best
    return FitResult(
        model=model,
        params=_expand(sol.x, free_index, fixed),
        weighted_sse=sse,
        n_bonds=len(bonds),
        n_starts=len(starts),
        converged=bool(sol.success),
        start_results=start_results,
    )
