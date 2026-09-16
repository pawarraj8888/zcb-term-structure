"""Load Wall Street Journal Treasury note and bond quotes.

The WSJ (Tullett Prebon) sheet quotes prices in 32nds: the two digits after
the decimal point are 32nds and the optional third digit is eighths of a 32nd.
So ``99.256`` means 99 + (25 + 6/8) / 32 = 99.8046875 per 100 face.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

import openpyxl
import pandas as pd

WSJ_DATE_FORMAT = "%A, %B %d, %Y"
HEADER_LABEL = "Maturity"
FOOTER_PREFIX = "Source:"
MAX_THIRTY_SECONDS = 31
MAX_EIGHTHS = 7


class QuoteFormatError(ValueError):
    """Raised when a WSJ quote cannot be interpreted."""


def parse_wsj_price(quote: float) -> float:
    """Convert a WSJ 32nds quote (e.g. 99.256) to a decimal price per 100 face."""
    if not isinstance(quote, (int, float)) or isinstance(quote, bool) or quote < 0:
        raise QuoteFormatError(f"Invalid WSJ price quote: {quote!r}")
    if abs(quote * 1000 - round(quote * 1000)) > 1e-6:
        raise QuoteFormatError(f"WSJ quotes have at most three decimals: {quote!r}")
    text = f"{quote:.3f}"
    whole_text, frac_text = text.split(".")
    thirty_seconds = int(frac_text[:2])
    eighths = int(frac_text[2])
    if thirty_seconds > MAX_THIRTY_SECONDS:
        raise QuoteFormatError(f"32nds digit out of range in quote {quote!r}")
    if eighths > MAX_EIGHTHS:
        raise QuoteFormatError(f"Eighths digit out of range in quote {quote!r}")
    return int(whole_text) + (thirty_seconds + eighths / 8.0) / 32.0


def parse_wsj_price_decimal_tenths(quote: float) -> float:
    """Alternative reading used only for the convention check: third digit = tenths of a 32nd."""
    text = f"{quote:.3f}"
    whole_text, frac_text = text.split(".")
    return int(whole_text) + (int(frac_text[:2]) + int(frac_text[2]) / 10.0) / 32.0


@dataclass(frozen=True)
class QuoteSheet:
    """Parsed WSJ quote sheet."""

    quote_date: date
    quotes: pd.DataFrame  # columns: maturity, coupon, bid_quote, ask_quote, chg, asked_yield


def _to_date(value) -> date:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    raise QuoteFormatError(f"Expected a date cell, got {value!r}")


def _parse_quote_date(rows: list[tuple]) -> date:
    for row in rows:
        for cell in row:
            if isinstance(cell, str):
                try:
                    return datetime.strptime(cell.strip(), WSJ_DATE_FORMAT).date()
                except ValueError:
                    continue
    raise QuoteFormatError("Quote date line (e.g. 'Friday, September 04, 2026') not found")


def load_wsj_quotes(path: str | Path) -> QuoteSheet:
    """Read the WSJ 'Treasury Notes & Bonds' workbook into a tidy DataFrame."""
    workbook = openpyxl.load_workbook(path, data_only=True, read_only=True)
    sheet = workbook.worksheets[0]
    rows = [tuple(r) for r in sheet.iter_rows(values_only=True)]
    workbook.close()

    quote_date = _parse_quote_date(rows)
    header_index = next(
        (i for i, row in enumerate(rows) if any(c == HEADER_LABEL for c in row)), None
    )
    if header_index is None:
        raise QuoteFormatError("Header row with 'Maturity' not found")
    header_row = rows[header_index]
    first_col = next(i for i, c in enumerate(header_row) if c == HEADER_LABEL)

    records = []
    for row in rows[header_index + 1 :]:
        cells = row[first_col : first_col + 6]
        if not cells or cells[0] is None:
            continue
        if isinstance(cells[0], str):
            if cells[0].startswith(FOOTER_PREFIX):
                break
            continue
        maturity, coupon, bid, ask, chg, asked_yield = cells
        records.append(
            {
                "maturity": _to_date(maturity),
                "coupon": float(coupon),
                "bid_quote": float(bid),
                "ask_quote": float(ask),
                "chg": chg,
                "asked_yield": float(asked_yield),
            }
        )
    if not records:
        raise QuoteFormatError("No bond rows found beneath the header")

    quotes = pd.DataFrame.from_records(records)
    quotes["bid_price"] = quotes["bid_quote"].map(parse_wsj_price)
    quotes["ask_price"] = quotes["ask_quote"].map(parse_wsj_price)
    quotes["mid_price"] = (quotes["bid_price"] + quotes["ask_price"]) / 2.0
    return QuoteSheet(quote_date=quote_date, quotes=quotes)
