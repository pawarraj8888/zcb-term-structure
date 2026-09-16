from datetime import date

import numpy as np
import pytest

from zcb import (
    check_conventions,
    fit_term_structure,
    load_wsj_quotes,
    payment_date_table,
    prepare_bonds,
    settlement_date_for,
)


def test_settlement_skips_weekend_and_labor_day():
    assert settlement_date_for(date(2026, 9, 4)) == date(2026, 9, 8)
    assert settlement_date_for(date(2026, 9, 9)) == date(2026, 9, 10)
    assert settlement_date_for(date(2026, 9, 4), business_days=2) == date(2026, 9, 9)


@pytest.fixture(scope="module")
def fitted(data_path):
    sheet = load_wsj_quotes(data_path)
    settlement = settlement_date_for(sheet.quote_date)
    bonds = prepare_bonds(sheet.quotes, settlement)
    return sheet, bonds, fit_term_structure(bonds, "svensson")


def test_recomputed_yields_match_wsj(fitted):
    _, bonds, _ = fitted
    diff_bp = (bonds.table["street_yield_pct"] - bonds.table["wsj_asked_yield"]) * 100
    assert np.sqrt(np.mean(diff_bp**2)) < 1.0
    assert diff_bp.abs().max() < 3.0


def test_convention_check_prefers_regular_way_and_eighths(data_path):
    sheet = load_wsj_quotes(data_path)
    best = check_conventions(sheet.quotes, sheet.quote_date).iloc[0]
    assert best["settlement"] == date(2026, 9, 8)
    assert best["price_format"] == "32nds + eighths"


def test_fit_quality(fitted):
    _, bonds, result = fitted
    assert result.summary["n_bonds_in_fit"] == 339
    assert result.summary["yield_rmse_bp"] < 4.0
    assert result.summary["converged"]
    assert result.bond_table["in_fit"].sum() == 339


def test_payment_date_table_covers_every_cash_flow(fitted):
    _, bonds, result = fitted
    table = payment_date_table(bonds, result.fit.params)
    assert table["payment_date"].is_monotonic_increasing
    assert (np.diff(table["discount_factor"]) < 0).all()
    assert table["payment_date"].iloc[0] == date(2026, 9, 15)
    assert table["payment_date"].iloc[-1] == date(2056, 8, 15)
    assert set(bonds.table["maturity"]).issubset(set(table["payment_date"]))
