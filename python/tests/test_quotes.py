from datetime import date

import pytest

from zcb.quotes import QuoteFormatError, load_wsj_quotes, parse_wsj_price


@pytest.mark.parametrize(
    "quote, expected",
    [
        (99.256, 99 + (25 + 6 / 8) / 32),  # 99-25 6/8 32nds
        (100.01, 100 + 1 / 32),            # 100-01
        (100.0, 100.0),
        (92.134, 92 + (13 + 4 / 8) / 32),
        (99.316, 99 + (31 + 6 / 8) / 32),
    ],
)
def test_parse_wsj_price_reads_32nds_and_eighths(quote, expected):
    assert parse_wsj_price(quote) == pytest.approx(expected, abs=1e-12)


def test_parse_wsj_price_rejects_invalid_32nds():
    with pytest.raises(QuoteFormatError):
        parse_wsj_price(99.326)  # 32/32 is not a valid tick


def test_parse_wsj_price_rejects_invalid_eighths():
    with pytest.raises(QuoteFormatError):
        parse_wsj_price(99.258)  # eighths digit must be 0-7


def test_parse_wsj_price_rejects_negative():
    with pytest.raises(QuoteFormatError):
        parse_wsj_price(-1.0)


def test_load_wsj_quotes_reads_all_bonds(data_path):
    sheet = load_wsj_quotes(data_path)
    assert sheet.quote_date == date(2026, 9, 4)
    assert len(sheet.quotes) == 353
    assert list(sheet.quotes.columns[:6]) == ["maturity", "coupon", "bid_quote", "ask_quote", "chg", "asked_yield"]
    assert (sheet.quotes["ask_price"] >= sheet.quotes["bid_price"]).all()
    first = sheet.quotes.iloc[0]
    assert first["maturity"] == date(2026, 9, 15)
    assert first["ask_price"] == pytest.approx(100 + 1 / 32)
