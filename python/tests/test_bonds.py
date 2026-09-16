import math
from datetime import date

import pytest

from zcb.bonds import (
    accrued_interest,
    add_months,
    build_schedule,
    continuous_yield,
    coupon_dates,
    dirty_price_from_yield,
    eay_to_continuous,
    modified_duration,
    street_yield,
    tbill_price,
)

SETTLE = date(2026, 9, 8)


def test_add_months_respects_month_end_rule():
    assert add_months(date(2027, 2, 28), -6, end_of_month=True) == date(2026, 8, 31)
    assert add_months(date(2027, 2, 15), -6, end_of_month=False) == date(2026, 8, 15)
    assert add_months(date(2026, 8, 31), 6, end_of_month=True) == date(2027, 2, 28)


def test_coupon_dates_for_thirty_year_bond():
    previous, future = coupon_dates(date(2056, 8, 15), SETTLE)
    assert previous == date(2026, 8, 15)
    assert future[0] == date(2027, 2, 15)
    assert future[-1] == date(2056, 8, 15)
    assert len(future) == 60


def test_coupon_dates_month_end_bond():
    previous, future = coupon_dates(date(2027, 2, 28), SETTLE)
    assert previous == date(2026, 8, 31)
    assert future == (date(2027, 2, 28),)


def test_coupon_dates_rejects_matured_bond():
    with pytest.raises(ValueError):
        coupon_dates(date(2026, 9, 8), SETTLE)


def test_schedule_amounts_and_times():
    # Settlement 8 Sep 2026: the 15 Sep 2026 coupon is still ahead, so three payments remain.
    schedule = build_schedule(date(2027, 9, 15), 4.0, SETTLE)
    assert schedule.dates == (date(2026, 9, 15), date(2027, 3, 15), date(2027, 9, 15))
    assert schedule.amounts == (2.0, 2.0, 102.0)
    assert schedule.previous_coupon == date(2026, 3, 15)
    assert schedule.times[-1] == pytest.approx((date(2027, 9, 15) - SETTLE).days / 365)


def test_accrued_interest_actual_actual():
    # 4.625% 15 Sep 2026: previous coupon 15 Mar 2026, 184-day period, 177 days accrued.
    schedule = build_schedule(date(2026, 9, 15), 4.625, SETTLE)
    assert accrued_interest(4.625, SETTLE, schedule) == pytest.approx(2.3125 * 177 / 184)


def test_street_yield_round_trips_price():
    schedule = build_schedule(date(2036, 2, 15), 4.5, SETTLE)
    dirty = dirty_price_from_yield(0.0467, SETTLE, schedule)
    assert street_yield(dirty, SETTLE, schedule) == pytest.approx(0.0467, abs=1e-12)


def test_street_yield_uses_simple_interest_in_final_period():
    schedule = build_schedule(date(2026, 9, 30), 0.875, SETTLE)
    accrued = accrued_interest(0.875, SETTLE, schedule)
    dirty = 99 + (26 + 6 / 8) / 32 + accrued
    w = (date(2026, 9, 30) - SETTLE).days / (date(2026, 9, 30) - date(2026, 3, 31)).days
    expected = (100.4375 / dirty - 1) * 2 / w
    assert street_yield(dirty, SETTLE, schedule) == pytest.approx(expected)
    assert dirty_price_from_yield(expected, SETTLE, schedule) == pytest.approx(dirty)


def test_modified_duration_is_positive_and_below_maturity():
    schedule = build_schedule(date(2056, 8, 15), 5.125, SETTLE)
    duration = modified_duration(0.0524, SETTLE, schedule)
    assert 10 < duration < 30


def test_lecture_tbill_example_slide_29():
    price = tbill_price(10_000, 0.05265, 43)
    assert price == pytest.approx(9937.11, abs=0.01)
    assert continuous_yield(price, 10_000, 43) == pytest.approx(0.05355, abs=5e-5)


def test_lecture_strip_example_slide_30():
    # Aug 15 2023 stripped principal, asked 89.482 on Sep 13 2017 (2161 days): EAY 1.89%, r 1.88%.
    r = continuous_yield(89.482, 100.0, 2161)
    assert r == pytest.approx(0.0188, abs=5e-5)
    eay = math.exp(r) - 1
    assert eay == pytest.approx(0.0189, abs=5e-5)
    assert eay_to_continuous(eay) == pytest.approx(r)
