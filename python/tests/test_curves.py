import numpy as np
import pytest

from zcb.curves import (
    BondPricingInput,
    discount_factor,
    fit_curve,
    svensson_forward,
    svensson_zero,
    weighted_residuals,
)

PARAMS = np.array([0.045, -0.010, 0.020, 0.030, 1.5, 12.0])


def test_zero_rate_limits():
    assert svensson_zero(0.0, *PARAMS) == pytest.approx(PARAMS[0] + PARAMS[1])
    assert svensson_zero(1e6, *PARAMS) == pytest.approx(PARAMS[0], abs=1e-6)


def test_forward_is_derivative_of_r_times_t():
    t = np.linspace(0.1, 30, 200)
    h = 1e-5
    numeric = (svensson_zero(t + h, *PARAMS) * (t + h) - svensson_zero(t - h, *PARAMS) * (t - h)) / (2 * h)
    np.testing.assert_allclose(svensson_forward(t, *PARAMS), numeric, atol=1e-7)


def test_discount_factor_matches_definition():
    t = np.array([0.5, 2.0, 10.0])
    np.testing.assert_allclose(discount_factor(t, PARAMS), np.exp(-svensson_zero(t, *PARAMS) * t))


def _synthetic_bonds(params: np.ndarray) -> list[BondPricingInput]:
    bonds = []
    rng = np.random.default_rng(7)
    for maturity in np.arange(0.5, 30.5, 0.5):
        times = np.arange(maturity % 0.5 or 0.5, maturity + 1e-9, 0.5)
        coupon = float(rng.uniform(1.0, 6.0))
        amounts = np.full(times.shape, coupon / 2)
        amounts[-1] += 100.0
        price = float(np.dot(amounts, discount_factor(times, params)))
        duration = float(np.dot(times, amounts * discount_factor(times, params)) / price)
        bonds.append(BondPricingInput(f"{maturity:.1f}y", times, amounts, price, 1.0 / duration))
    return bonds


def test_fit_recovers_synthetic_svensson_curve():
    bonds = _synthetic_bonds(PARAMS)
    result = fit_curve(bonds, "svensson")
    grid = np.linspace(0.1, 30, 50)
    np.testing.assert_allclose(svensson_zero(grid, *result.params), svensson_zero(grid, *PARAMS), atol=2e-5)
    assert result.weighted_sse < 1e-8
    assert np.max(np.abs(weighted_residuals(result.params, bonds))) < 1e-4


def test_nelson_siegel_pins_beta3():
    bonds = _synthetic_bonds(PARAMS)
    result = fit_curve(bonds, "nelson_siegel")
    assert result.params[3] == 0.0
    assert result.weighted_sse >= 0.0


def test_fixed_parameter_is_respected():
    bonds = _synthetic_bonds(PARAMS)
    result = fit_curve(bonds, "svensson", fixed_params={"beta0": 0.03})
    assert result.params[0] == 0.03


def test_unknown_model_rejected():
    with pytest.raises(ValueError):
        fit_curve(_synthetic_bonds(PARAMS), "cubic")
