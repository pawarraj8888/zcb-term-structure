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

FIRST_ROW = 4
CHECKS = [
    # (Bonds column, results key or callable, tolerance, label)
    ("G", lambda b: b["clean_price"], 1e-9, "ask clean price (32nds parse)"),
    ("L", lambda b: b["accrued_interest"], 1e-9, "accrued interest"),
    ("M", lambda b: b["dirty_price"], 1e-9, "dirty price"),
    ("N", lambda b: b["street_yield_pct"], 1e-7, "street yield % (Excel YIELD)"),
    ("P", lambda b: b["modified_duration"], 1e-7, "modified duration (Excel MDURATION)"),
    ("Q", lambda b: b["n_cash_flows"], 0, "remaining cash flows (COUPNUM)"),
    ("S", lambda b: 1 if b["in_fit"] else 0, 0, "in-fit flag"),
    ("T", lambda b: b["model_clean_price"] + b["accrued_interest"], 1e-7, "model dirty price"),
    ("U", lambda b: b["model_clean_price"], 1e-7, "model clean price"),
    ("W", lambda b: b["model_yield_pct"], 1e-6, "model yield %"),
    ("X", lambda b: b["yield_error_bp"], 1e-4, "yield error (bp)"),
]


def compare(workbook: Path, results: Path) -> int:
    res = json.loads(results.read_text())
    wb = openpyxl.load_workbook(workbook, data_only=True)
    bonds = res["bonds"]
    ws = wb["Bonds"]
    failures = 0
    print(f"{'check':45s} {'max |Excel - Python|':>22s}  tolerance")
    for col, getter, tol, label in CHECKS:
        excel = np.array([ws[f"{col}{FIRST_ROW + i}"].value for i in range(len(bonds))], dtype=float)
        python = np.array([getter(b) for b in bonds], dtype=float)
        if np.isnan(excel).any():
            print(f"{label:45s} {'MISSING (not recalculated?)':>22s}")
            failures += 1
            continue
        diff = float(np.max(np.abs(excel - python)))
        ok = diff <= tol
        failures += 0 if ok else 1
        print(f"{label:45s} {diff:22.3e}  {tol:g} {'OK' if ok else 'FAIL'}")

    inputs = wb["Inputs"]
    sv = res["svensson"]
    settle = inputs["B6"].value
    settle = settle.date() if isinstance(settle, datetime) else settle
    scalars = [
        ("settlement date", settle == date.fromisoformat(res["settlement"]), None),
        ("weighted SSE", inputs["B21"].value, sv["weighted_sse"]),
        ("bonds in fit", inputs["B22"].value, sv["n_bonds_in_fit"]),
        ("price RMSE", inputs["B24"].value, sv["price_rmse"]),
        ("yield RMSE (bp)", inputs["B25"].value, sv["yield_rmse_bp"]),
        ("yield MAE (bp)", inputs["B26"].value, sv["yield_mae_bp"]),
        ("max |yield error| (bp)", inputs["B27"].value, sv["yield_max_abs_bp"]),
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

    zc = wb["ZeroCurve"]
    pay = res["payment_dates"]
    print()
    for col, key, label in (("D", "zero_rate_cc_pct", "zero rate %"), ("E", "discount_factor", "discount factor"), ("F", "inst_forward_cc_pct", "forward %")):
        excel = np.array([zc[f"{col}{4 + i}"].value for i in range(len(pay))], dtype=float)
        python = np.array([p[key] for p in pay], dtype=float)
        diff = float(np.max(np.abs(excel - python)))
        ok = diff <= 1e-8
        failures += 0 if ok else 1
        print(f"{'ZeroCurve ' + label:45s} {diff:22.3e}  {'OK' if ok else 'FAIL'}  ({len(pay)} payment dates)")
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
