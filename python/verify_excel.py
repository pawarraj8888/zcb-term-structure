#!/usr/bin/env python3
"""Compare the Excel workbook's calculated values (after a recalculation in
Excel) with the Python results. Exit code 1 if any difference exceeds tolerance."""
from __future__ import annotations

import argparse
import json
import sys
from datetime import date, datetime
from pathlib import Path

import numpy as np
import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parent))
from build_excel import BOND_COLS, BOND_FIRST_ROW, CURVE_SSE_CELL, CURVE_STATS, CURVE_TABLE_FIRST_ROW  # noqa: E402

CHECKS = [
    # (Bonds column key, python getter, tolerance, label)
    ("clean", lambda b: b["clean_price"], 1e-9, "asked clean price (32nds parse)"),
    ("accrued", lambda b: b["accrued_interest"], 1e-9, "accrued interest"),
    ("dirty", lambda b: b["dirty_price"], 1e-9, "dirty price"),
    ("ytm", lambda b: b["street_yield_pct"], 1e-7, "street yield % (Excel YIELD)"),
    ("duration", lambda b: b["modified_duration"], 1e-7, "modified duration (Excel MDURATION)"),
    ("n_coupons", lambda b: b["n_cash_flows"], 0, "remaining cash flows (COUPNUM)"),
    ("use", lambda b: 1 if b["in_fit"] else 0, 0, "in-fit flag"),
    ("model_dirty", lambda b: b["model_clean_price"] + b["accrued_interest"], 1e-7, "model dirty price"),
    ("model_clean", lambda b: b["model_clean_price"], 1e-7, "model clean price"),
    ("model_ytm", lambda b: b["model_yield_pct"], 1e-6, "model yield %"),
    ("yield_err", lambda b: b["yield_error_bp"], 1e-4, "yield error (bp)"),
]


def compare(workbook: Path, results: Path) -> int:
    res = json.loads(results.read_text())
    wb = openpyxl.load_workbook(workbook, data_only=True)
    bonds = res["bonds"]
    ws = wb["Bonds"]
    failures = 0
    print(f"{'check':45s} {'max |Excel - Python|':>22s}  tolerance")
    for key, getter, tol, label in CHECKS:
        col = BOND_COLS[key]
        raw = [ws[f"{col}{BOND_FIRST_ROW + i}"].value for i in range(len(bonds))]
        if any(v is None or isinstance(v, str) for v in raw):
            bad = next(v for v in raw if v is None or isinstance(v, str))
            print(f"{label:45s} {'NOT NUMERIC: ' + repr(bad):>22s}  FAIL")
            failures += 1
            continue
        diff = float(np.max(np.abs(np.array(raw, dtype=float) - np.array([getter(b) for b in bonds], dtype=float))))
        ok = diff <= tol
        failures += 0 if ok else 1
        print(f"{label:45s} {diff:22.3e}  {tol:g} {'OK' if ok else 'FAIL'}")

    curve = wb["Curve"]
    sv = res["svensson"]
    settle = wb["Bonds"]["B3"].value
    settle = settle.date() if isinstance(settle, datetime) else settle
    scalars = [
        ("settlement date", settle == date.fromisoformat(res["settlement"]), None),
        ("weighted SSE", curve[CURVE_SSE_CELL].value, sv["weighted_sse"]),
        ("bonds in fit", curve[CURVE_STATS["n_used"]].value, sv["n_bonds_in_fit"]),
        ("price RMSE", curve[CURVE_STATS["price_rmse"]].value, sv["price_rmse"]),
        ("yield RMSE (bp)", curve[CURVE_STATS["yield_rmse"]].value, sv["yield_rmse_bp"]),
        ("yield MAE (bp)", curve[CURVE_STATS["yield_mae"]].value, sv["yield_mae_bp"]),
        ("max |yield error| (bp)", curve[CURVE_STATS["yield_max"]].value, sv["yield_max_abs_bp"]),
    ]
    print()
    for label, excel_value, python_value in scalars:
        if python_value is None:
            ok = bool(excel_value)
            print(f"{label:45s} {'matches' if ok else 'MISMATCH':>22s}")
        else:
            diff = abs(float(excel_value) - float(python_value))
            ok = diff <= 1e-6 * max(1.0, abs(float(python_value)))
            print(f"{label:45s} Excel {float(excel_value):.10g} | Python {float(python_value):.10g} {'OK' if ok else 'FAIL'}")
        failures += 0 if ok else 1

    pay = res["payment_dates"]
    print()
    for col, key, label in (("D", "zero_rate_cc_pct", "zero rate %"), ("E", "discount_factor", "discount factor"),
                            ("F", "inst_forward_cc_pct", "forward %")):
        excel = np.array([curve[f"{col}{CURVE_TABLE_FIRST_ROW + i}"].value for i in range(len(pay))], dtype=float)
        python = np.array([p[key] for p in pay], dtype=float)
        diff = float(np.max(np.abs(excel - python)))
        ok = diff <= 1e-8
        failures += 0 if ok else 1
        print(f"{'Curve table ' + label:45s} {diff:22.3e}  {'OK' if ok else 'FAIL'}  ({len(pay)} payment dates)")
    print(f"\n{'ALL CHECKS PASSED' if failures == 0 else f'{failures} CHECK(S) FAILED'}")
    return 0 if failures == 0 else 1


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--workbook", default="../excel/Miniproject2_ZCB_Term_Structure.xlsx")
    parser.add_argument("--results", default="../output/results.json")
    args = parser.parse_args(argv)
    return compare(Path(args.workbook), Path(args.results))


if __name__ == "__main__":
    sys.exit(main())
