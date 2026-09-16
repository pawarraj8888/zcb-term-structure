"""US Treasury note/bond mechanics: coupon schedules, accrued interest,
street-convention yield, and duration.

Conventions (matching Excel's COUP*/YIELD/MDURATION functions with basis 1):
* semi-annual coupons on the maturity day-of-month; end-of-month maturities pay
  on month-ends (a 28 Feb maturity pays 31 Aug / 28 Feb);
* accrued interest is actual/actual within the coupon period;
* yield is the Treasury street convention: semi-annual compounding with a
  fractional first period, and simple interest when one cash flow remains.
"""
from __future__ import annotations

import calendar
import math
from dataclasses import dataclass
from datetime import date

import numpy as np
from scipy.optimize import brentq

FACE_VALUE = 100.0
COUPONS_PER_YEAR = 2
MONTHS_PER_COUPON = 12 // COUPONS_PER_YEAR
DAYS_PER_YEAR = 365.0
YIELD_LOWER_BOUND = -0.5
YIELD_UPPER_BOUND = 2.0


def is_month_end(day: date) -> bool:
    return day.day == calendar.monthrange(day.year, day.month)[1]


def add_months(day: date, months: int, end_of_month: bool) -> date:
    """Shift a date by whole months, preserving the day-of-month or month-end rule."""
    month_index = day.year * 12 + (day.month - 1) + months
    year, month = divmod(month_index, 12)
    month += 1
    last_day = calendar.monthrange(year, month)[1]
    target_day = last_day if end_of_month else min(day.day, last_day)
    return date(year, month, target_day)


@dataclass(frozen=True)
class CashFlowSchedule:
    """Remaining cash flows of a bond as seen from the settlement date."""

    previous_coupon: date
    dates: tuple[date, ...]
    amounts: tuple[float, ...]
    times: tuple[float, ...]  # years from settlement (actual/365)

    @property
    def next_coupon(self) -> date:
        return self.dates[0]


def coupon_dates(maturity: date, settlement: date) -> tuple[date, tuple[date, ...]]:
    """Return (previous coupon date <= settlement, future coupon dates > settlement)."""
    if maturity <= settlement:
        raise ValueError(f"Maturity {maturity} is not after settlement {settlement}")
    eom = is_month_end(maturity)
    future: list[date] = []
    step = 0
    current = maturity
    while current > settlement:
        future.append(current)
        step += 1
        current = add_months(maturity, -MONTHS_PER_COUPON * step, eom)
    return current, tuple(reversed(future))


def build_schedule(maturity: date, coupon_rate: float, settlement: date) -> CashFlowSchedule:
    """Cash flows per 100 face: coupon/2 on each coupon date, plus face at maturity."""
    previous, future = coupon_dates(maturity, settlement)
    periodic_coupon = coupon_rate / 100.0 * FACE_VALUE / COUPONS_PER_YEAR
    amounts = tuple(
        periodic_coupon + (FACE_VALUE if d == maturity else 0.0) for d in future
    )
    times = tuple((d - settlement).days / DAYS_PER_YEAR for d in future)
    return CashFlowSchedule(previous, future, amounts, times)


def accrued_interest(coupon_rate: float, settlement: date, schedule: CashFlowSchedule) -> float:
    """Actual/actual accrued interest per 100 face."""
    period_days = (schedule.next_coupon - schedule.previous_coupon).days
    accrued_days = (settlement - schedule.previous_coupon).days
    periodic_coupon = coupon_rate / 100.0 * FACE_VALUE / COUPONS_PER_YEAR
    return periodic_coupon * accrued_days / period_days


def fraction_of_period_remaining(settlement: date, schedule: CashFlowSchedule) -> float:
    period_days = (schedule.next_coupon - schedule.previous_coupon).days
    return (schedule.next_coupon - settlement).days / period_days


def dirty_price_from_yield(yield_rate: float, settlement: date, schedule: CashFlowSchedule) -> float:
    """Street-convention dirty price for a given bond-equivalent yield."""
    w = fraction_of_period_remaining(settlement, schedule)
    amounts = np.asarray(schedule.amounts)
    if len(amounts) == 1:
        # Simple interest in the final coupon period (Excel YIELD / Treasury convention).
        return amounts[0] / (1.0 + yield_rate * w / COUPONS_PER_YEAR)
    exponents = np.arange(len(amounts)) + w
    discount = (1.0 + yield_rate / COUPONS_PER_YEAR) ** (-exponents)
    return float(np.dot(amounts, discount))


def street_yield(dirty_price: float, settlement: date, schedule: CashFlowSchedule) -> float:
    """Bond-equivalent yield (semi-annual) that reproduces the dirty price."""
    w = fraction_of_period_remaining(settlement, schedule)
    if len(schedule.amounts) == 1:
        return (schedule.amounts[0] / dirty_price - 1.0) * COUPONS_PER_YEAR / w

    def objective(y: float) -> float:
        return dirty_price_from_yield(y, settlement, schedule) - dirty_price

    return brentq(objective, YIELD_LOWER_BOUND, YIELD_UPPER_BOUND, xtol=1e-14, maxiter=200)


def modified_duration(yield_rate: float, settlement: date, schedule: CashFlowSchedule) -> float:
    """Modified duration in years, Excel MDURATION style (compounded in every period)."""
    w = fraction_of_period_remaining(settlement, schedule)
    amounts = np.asarray(schedule.amounts)
    exponents = np.arange(len(amounts)) + w
    discount = (1.0 + yield_rate / COUPONS_PER_YEAR) ** (-exponents)
    present_values = amounts * discount
    macaulay_years = float(np.dot(exponents / COUPONS_PER_YEAR, present_values) / present_values.sum())
    return macaulay_years / (1.0 + yield_rate / COUPONS_PER_YEAR)


# --- Conversions from the lecture (slides 29-30), used for illustration/tests ---
TBILL_DAY_BASIS = 360.0


def tbill_price(face: float, quoted_discount_yield: float, days: int) -> float:
    """T-bill price from the quoted (bank discount) yield: F (1 - d * days / 360)."""
    return face * (1.0 - quoted_discount_yield * days / TBILL_DAY_BASIS)


def continuous_yield(price: float, face: float, days: int) -> float:
    """Continuously compounded ZCB yield: -ln(P/F) / (days / 365)."""
    return -math.log(price / face) / (days / DAYS_PER_YEAR)


def eay_to_continuous(effective_annual_yield: float) -> float:
    """r = ln(1 + EAY)."""
    return math.log1p(effective_annual_yield)
