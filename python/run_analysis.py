#!/usr/bin/env python3
"""Run the full Miniproject 2 pipeline: parse WSJ quotes, fit the Svensson zero
curve, and write tables, JSON and figures.

    python run_analysis.py --data ../data/Treasury_data_090426.xlsx --out ../output --site ../docs
"""
from __future__ import annotations

import argparse
import sys
from datetime import date
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))

from zcb import (  # noqa: E402
    check_conventions,
    fit_term_structure,
    load_wsj_quotes,
    payment_date_table,
    prepare_bonds,
    profile_beta0,
    settlement_date_for,
)
from zcb.analysis import DEFAULT_MIN_YEARS_IN_FIT  # noqa: E402
from zcb.report import (  # noqa: E402
    frame_records,
    plot_discount_factors,
    plot_residuals,
    plot_zero_curve,
    write_json,
)

BOND_TABLE_COLUMNS = [
    "label", "maturity", "coupon", "bid_quote", "ask_quote", "clean_price", "accrued_interest",
    "dirty_price", "wsj_asked_yield", "street_yield_pct", "modified_duration", "years_to_maturity",
    "n_cash_flows", "in_fit", "model_clean_price", "price_error", "model_yield_pct", "yield_error_bp",
]


def parse_args(argv=None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--data", default="../data/Treasury_data_090426.xlsx", help="WSJ quote workbook")
    parser.add_argument("--out", default="../output", help="directory for CSV/JSON/PNG outputs")
    parser.add_argument("--site", default=None, help="optional docs/ directory to receive data.json")
    parser.add_argument("--settlement", default=None, help="override settlement date (YYYY-MM-DD)")
    parser.add_argument("--price-side", default="ask", choices=["ask", "bid", "mid"])
    parser.add_argument("--min-years", type=float, default=DEFAULT_MIN_YEARS_IN_FIT,
                        help="exclude bonds with less time to maturity from the fit")
    return parser.parse_args(argv)


def main(argv=None) -> int:
    args = parse_args(argv)
    out = Path(args.out)
    figures = out / "figures"
    figures.mkdir(parents=True, exist_ok=True)

    sheet = load_wsj_quotes(args.data)
    settlement = date.fromisoformat(args.settlement) if args.settlement else settlement_date_for(sheet.quote_date)
    print(f"Quotes: {len(sheet.quotes)} Treasury notes/bonds as of {sheet.quote_date}; settlement {settlement}")

    conventions = check_conventions(sheet.quotes, sheet.quote_date)
    print("\nConvention check (our asked yield vs WSJ asked yield):")
    print(conventions.to_string(index=False))

    bonds = prepare_bonds(sheet.quotes, settlement, args.price_side, args.min_years)
    svensson = fit_term_structure(bonds, "svensson")
    nelson_siegel = fit_term_structure(bonds, "nelson_siegel")
    payments = payment_date_table(bonds, svensson.fit.params)
    beta0_profile = profile_beta0(bonds)

    print("\nSvensson fit:")
    for key, value in svensson.summary.items():
        print(f"  {key}: {value}")
    print(f"\nNelson-Siegel benchmark: yield RMSE {nelson_siegel.summary['yield_rmse_bp']:.2f} bp "
          f"vs Svensson {svensson.summary['yield_rmse_bp']:.2f} bp")
    print(f"\nPayment dates covered: {len(payments)} (first {payments.payment_date.iloc[0]}, last {payments.payment_date.iloc[-1]})")

    payments.to_csv(out / "zero_curve_payment_dates.csv", index=False, float_format="%.10g")
    svensson.bond_table[BOND_TABLE_COLUMNS].to_csv(out / "bond_fit.csv", index=False, float_format="%.10g")
    conventions.to_csv(out / "convention_check.csv", index=False, float_format="%.6g")
    beta0_profile.to_csv(out / "beta0_profile.csv", index=False, float_format="%.8g")
    write_json({"quote_date": sheet.quote_date, "settlement": settlement, **svensson.summary}, out / "parameters.json")

    plot_zero_curve(svensson.bond_table, svensson.fit.params, sheet.quote_date, settlement, figures / "zero_curve.png")
    plot_residuals(svensson.bond_table, svensson.summary, figures / "yield_residuals.png")
    plot_discount_factors(payments, figures / "discount_factors.png")

    payload = {
        "quote_date": sheet.quote_date,
        "settlement": settlement,
        "svensson": svensson.summary,
        "nelson_siegel": nelson_siegel.summary,
        "conventions": frame_records(conventions),
        "beta0_profile": frame_records(beta0_profile),
        "payment_dates": frame_records(payments),
        "bonds": frame_records(svensson.bond_table[BOND_TABLE_COLUMNS]),
    }
    write_json(payload, out / "results.json")
    if args.site:
        site = Path(args.site)
        site.mkdir(parents=True, exist_ok=True)
        write_json(payload, site / "data.json")
    print(f"\nWrote outputs to {out.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
