#!/usr/bin/env python3
"""Create and execute the Jupyter notebook version of the analysis."""
from __future__ import annotations

import argparse
from pathlib import Path

import nbformat
from nbclient import NotebookClient
from nbformat.v4 import new_code_cell, new_markdown_cell, new_notebook

CELLS = [
    ("md", """# Miniproject 2 - Creating a ZCB Term Structure

**FRE 6103 Valuation for Financial Engineering (NYU Tandon)** | Raj Pawar

Fit a continuously compounded zero-coupon (ZCB) term structure to all US Treasury note and bond quotes
(WSJ, Sep 4, 2026) with the Svensson (1994) functional form, and report the discount rate for every
Treasury payment date. The reusable code lives in the `zcb` package next to this notebook
(`python/zcb/`); this notebook walks through the pipeline and shows the results."""),
    ("code", """import sys
from pathlib import Path

import numpy as np
import pandas as pd
from IPython.display import Image, display

sys.path.insert(0, str(Path("../python").resolve()))
from zcb import (load_wsj_quotes, settlement_date_for, check_conventions, prepare_bonds,
                 fit_term_structure, payment_date_table, profile_beta0)
from zcb.report import plot_zero_curve, plot_residuals, plot_discount_factors

pd.set_option("display.width", 160)
pd.set_option("display.max_columns", 30)
import tempfile
DATA = Path("../data/Treasury_data_090426.xlsx")
OUT = Path(tempfile.mkdtemp(prefix="zcb_notebook_"))  # scratch dir: the pipeline's output/ folder is written by run_analysis.py
(OUT / "figures").mkdir(parents=True, exist_ok=True)"""),
    ("md", """## 1. Data: WSJ Treasury notes and bonds

Prices are quoted in 32nds; the third decimal is eighths of a 32nd (`99.256` = 99 + 25.75/32).
Regular-way settlement is T+1 business days: the quote is Friday Sep 4, 2026, Monday Sep 7 is Labor Day,
so settlement is Tuesday Sep 8, 2026."""),
    ("code", """sheet = load_wsj_quotes(DATA)
settlement = settlement_date_for(sheet.quote_date)
print(f"Quote date {sheet.quote_date} | settlement {settlement} | {len(sheet.quotes)} bonds")
sheet.quotes.head(8)"""),
    ("md", """### Convention check
We recompute each bond's asked yield under alternative settlement rules and price readings and compare
with the WSJ *Asked yield* column. Regular-way settlement with the 32nds+eighths reading reproduces the
WSJ yields to a median of a few hundredths of a basis point."""),
    ("code", """conventions = check_conventions(sheet.quotes, sheet.quote_date)
conventions.round(3)"""),
    ("md", """## 2. Bond mechanics
Coupon schedule (month-end rule), actual/actual accrued interest, dirty price, Treasury street yield
and modified duration for every bond. Bonds with under 3 months to maturity are flagged `in_fit = False`
(money-market segment, following Gurkaynak-Sack-Wright 2007) and kept as an out-of-sample check."""),
    ("code", """bonds = prepare_bonds(sheet.quotes, settlement, price_side="ask")
cols = ["label", "clean_price", "accrued_interest", "dirty_price", "street_yield_pct", "wsj_asked_yield",
        "modified_duration", "n_cash_flows", "in_fit"]
print(f"{bonds.table.in_fit.sum()} of {len(bonds.table)} bonds enter the fit")
bonds.table[cols].iloc[[0, 5, 50, 150, 250, 352]].round(4)"""),
    ("md", r"""## 3. Model and estimation

Continuously compounded zero rate (Svensson 1994):

$$r(t)=\beta_0+\beta_1\frac{1-e^{-t/\tau_1}}{t/\tau_1}+\beta_2\left(\frac{1-e^{-t/\tau_1}}{t/\tau_1}-e^{-t/\tau_1}\right)+\beta_3\left(\frac{1-e^{-t/\tau_2}}{t/\tau_2}-e^{-t/\tau_2}\right)$$

Discount factor $d(t)=e^{-r(t)t}$; model dirty price of bond $i$ is $\sum_k CF_{ik}\,d(t_{ik})$.
Parameters minimise $\sum_i \big[(P_i^{model}-P_i^{market})/(P_i^{market} D_i)\big]^2$: each price error divided by
dirty price times modified duration is the first-order yield error, so this is the sum of squared yield errors.
The fit runs from 19 starting points on a $(\tau_1,\tau_2)$ grid, with
$\beta_0\ge 0$. Nelson-Siegel ($\beta_3=0$) is fitted as a benchmark."""),
    ("code", """svensson = fit_term_structure(bonds, "svensson")
nelson_siegel = fit_term_structure(bonds, "nelson_siegel")
summary = pd.DataFrame([svensson.summary, nelson_siegel.summary]).set_index("model")
summary[["n_bonds_in_fit", "weighted_sse", "price_rmse", "yield_rmse_bp", "yield_mae_bp", "yield_max_abs_bp",
         "n_starts", "n_starts_at_optimum"]].round(4)"""),
    ("code", """pd.Series(svensson.summary["params"], name="Svensson parameters").to_frame().round(6)"""),
    ("md", """## 4. Deliverable: discount rate for every Treasury payment date"""),
    ("code", """payments = payment_date_table(bonds, svensson.fit.params)
print(f"{len(payments)} distinct payment dates, {payments.payment_date.iloc[0]} to {payments.payment_date.iloc[-1]}")
pd.concat([payments.head(10), payments.tail(5)]).round(6)"""),
    ("code", """plot_zero_curve(svensson.bond_table, svensson.fit.params, sheet.quote_date, settlement, OUT / "figures/zero_curve.png")
display(Image(OUT / "figures/zero_curve.png", width=900))"""),
    ("code", """plot_discount_factors(payments, OUT / "figures/discount_factors.png")
display(Image(OUT / "figures/discount_factors.png", width=900))"""),
    ("md", """## 5. Fit quality"""),
    ("code", """plot_residuals(svensson.bond_table, svensson.summary, OUT / "figures/yield_residuals.png")
display(Image(OUT / "figures/yield_residuals.png", width=900))
tbl = svensson.bond_table
tbl["bucket"] = pd.cut(tbl.years_to_maturity, [0, 0.25, 1, 2, 3, 5, 7, 10, 20, 31])
tbl[tbl.in_fit].groupby("bucket", observed=True).yield_error_bp.agg(
    n="count", mean="mean", std="std", max_abs=lambda s: s.abs().max()).round(2)"""),
    ("code", """print("Largest yield errors among fitted bonds:")
tbl[tbl.in_fit].reindex(tbl[tbl.in_fit].yield_error_bp.abs().sort_values(ascending=False).index)[
    ["label", "years_to_maturity", "clean_price", "street_yield_pct", "model_yield_pct", "yield_error_bp", "price_error"]].head(6).round(4)"""),
    ("code", """print("Out-of-sample: bonds under 3 months excluded from the fit")
tbl[~tbl.in_fit][["label", "years_to_maturity", "street_yield_pct", "model_yield_pct", "yield_error_bp", "price_error"]].round(4)"""),
    ("md", """### Robustness: how well is the long-run level identified?
The asymptote $\\beta_0$ sits at its non-negativity bound. Pinning it anywhere from 0% to 3% barely changes the
objective or the curve inside the 30-year sample, so payment-date discount factors are robust to it."""),
    ("code", """profile_beta0(bonds).round(4)"""),
    ("md", """## 6. Methodology (one page)
See `writeup/Miniproject2_ZCB_Term_Structure_Writeup.md` (also as PDF) and the project page at
https://pawarraj8888.github.io/zcb-term-structure/ for the interactive chart and full tables.
The Excel implementation (`excel/Miniproject2_ZCB_Term_Structure.xlsx`) reproduces these numbers with live formulas."""),
]


def build(path: Path, execute: bool) -> None:
    nb = new_notebook()
    nb.metadata["kernelspec"] = {"name": "python3", "display_name": "Python 3", "language": "python"}
    nb.metadata["language_info"] = {"name": "python"}
    nb.cells = [new_markdown_cell(src) if kind == "md" else new_code_cell(src) for kind, src in CELLS]
    if execute:
        client = NotebookClient(nb, timeout=600, kernel_name="python3", resources={"metadata": {"path": str(path.parent)}})
        client.execute()
    path.parent.mkdir(parents=True, exist_ok=True)
    nbformat.write(nb, path)


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--out", default="../notebooks/Miniproject2_ZCB_Term_Structure.ipynb")
    parser.add_argument("--no-execute", action="store_true")
    args = parser.parse_args(argv)
    build(Path(args.out), execute=not args.no_execute)
    print(f"Notebook written to {Path(args.out).resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
