#!/usr/bin/env python3
"""Build the Excel implementation of Miniproject 2 with live formulas.

The workbook re-derives everything from the raw WSJ quotes with Excel's own
bond functions (COUPPCD/COUPNCD/COUPDAYBS/COUPDAYS/YIELD/MDURATION) and prices
every bond off the Svensson curve whose six parameters sit on the Inputs sheet.
Solver can be run on the SSE cell; the parameters are pre-loaded with the
Python optimum so the workbook opens at the fitted solution.

    python build_excel.py --data ../data/Treasury_data_090426.xlsx \
        --results ../output/results.json --out ../excel/Miniproject2_ZCB_Term_Structure.xlsx
"""
from __future__ import annotations

import argparse
import json
import sys
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import Reference, ScatterChart, Series
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName

sys.path.insert(0, str(Path(__file__).resolve().parent))
from zcb import load_wsj_quotes  # noqa: E402

TITLE = "Miniproject 2 - Creating a ZCB Term Structure"
COURSE = "FRE 6103 Valuation for Financial Engineering (NYU Tandon), Prof. David Shimko"
AUTHOR = "Raj Pawar"
LABOR_DAY_2026 = date(2026, 9, 7)

HEADER_FILL = PatternFill("solid", fgColor="1F3A5F")
INPUT_FILL = PatternFill("solid", fgColor="FFF4CE")
NOTE_FILL = PatternFill("solid", fgColor="EEF3FA")
HEADER_FONT = Font(bold=True, color="FFFFFF")
BOLD = Font(bold=True)
TITLE_FONT = Font(bold=True, size=16, color="1F3A5F")
THIN = Side(style="thin", color="C3C2B7")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical="top")

DATE_FMT = "yyyy-mm-dd"
PCT3 = "0.000"
PX = "0.0000"

FIRST_ROW = 4  # first bond row on Bonds / CashFlows

# Svensson zero rate and forward as Excel formula fragments with {t} substituted.
# Defined names for the parameters. They must not look like cell references
# (Excel would read "tau1" as column TAU, row 1), hence the Sv_ prefix.
NAME_OF = {"beta0": "Sv_beta0", "beta1": "Sv_beta1", "beta2": "Sv_beta2", "beta3": "Sv_beta3",
           "tau1": "Sv_tau1", "tau2": "Sv_tau2"}
ZERO_FORMULA = (
    "Sv_beta0+Sv_beta1*(1-EXP(-{t}/Sv_tau1))/({t}/Sv_tau1)"
    "+Sv_beta2*((1-EXP(-{t}/Sv_tau1))/({t}/Sv_tau1)-EXP(-{t}/Sv_tau1))"
    "+Sv_beta3*((1-EXP(-{t}/Sv_tau2))/({t}/Sv_tau2)-EXP(-{t}/Sv_tau2))"
)
FORWARD_FORMULA = (
    "Sv_beta0+Sv_beta1*EXP(-{t}/Sv_tau1)+Sv_beta2*({t}/Sv_tau1)*EXP(-{t}/Sv_tau1)"
    "+Sv_beta3*({t}/Sv_tau2)*EXP(-{t}/Sv_tau2)"
)


def header(ws, row: int, labels: list[str], start_col: int = 1) -> None:
    for offset, label in enumerate(labels):
        cell = ws.cell(row=row, column=start_col + offset, value=label)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        cell.border = BOX


def set_widths(ws, widths: dict[str, float]) -> None:
    for col, width in widths.items():
        ws.column_dimensions[col].width = width


def add_name(wb: Workbook, name: str, ref: str) -> None:
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


# ----------------------------------------------------------------------------- sheets
def build_cover(wb: Workbook, results: dict) -> None:
    ws = wb.active
    ws.title = "Cover"
    ws["B2"] = TITLE
    ws["B2"].font = TITLE_FONT
    ws["B3"] = COURSE
    ws["B4"] = f"{AUTHOR}  |  Data: WSJ / Tullett Prebon Treasury notes & bonds, {results['quote_date']}"
    lines = [
        ("What this workbook does",
         "Fits a continuously compounded zero-coupon (ZCB) term structure to all US Treasury note and bond "
         "quotes using the Svensson (1994) functional form, then reports the discount rate r(t) and discount "
         "factor d(t) = exp(-r(t) t) for every Treasury payment date (coupons and principal)."),
        ("Sheet map",
         "Inputs: dates, Svensson parameters, objective (weighted SSE) and fit statistics. Solver runs here.\n"
         "Data: raw WSJ quotes exactly as downloaded.\n"
         "Bonds: one row per bond - 32nds conversion, coupon dates (COUPPCD/COUPNCD), accrued interest, dirty price, street yield (Excel YIELD), "
         "duration, model price off the curve, price and yield errors.\n"
         "CashFlows: every remaining coupon/principal date of every bond, its zero rate and present value.\n"
         "ZeroCurve: the deliverable - r(t), d(t) and instantaneous forward f(t) at every payment date, with chart.\n"
         "Robustness: convention check, Nelson-Siegel benchmark and beta0 profile (values computed by the Python implementation).\n"
         "Methodology: the one-page write-up."),
        ("Live formulas",
         "Everything on Bonds, CashFlows and ZeroCurve is a formula. Change a parameter on Inputs and all "
         "model prices, errors and the curve update. Values on Data and Robustness are inputs/records."),
        ("Re-running the fit with Solver",
         "Data > Solver. Set Objective: Inputs!$B$21 (weighted SSE) to Min. By Changing: Inputs!$B$13:$B$18. "
         "Subject to: $B$13 >= 0, $B$17 >= 0.05, $B$18 >= 0.05. Method: GRG Nonlinear; untick 'Make Unconstrained "
         "Variables Non-Negative' (beta1..beta3 may be negative). The pre-loaded parameters are already the optimum "
         "found by the Python implementation (multi-start least squares), so Solver should not move them materially."),
        ("Companion Python implementation",
         "github.com/pawarraj8888/zcb-term-structure - identical conventions; the Excel SSE and every model price "
         "reproduce the Python numbers."),
    ]
    row = 6
    for heading, text in lines:
        ws.cell(row=row, column=2, value=heading).font = BOLD
        cell = ws.cell(row=row + 1, column=2, value=text)
        cell.alignment = WRAP
        ws.merge_cells(start_row=row + 1, start_column=2, end_row=row + 1, end_column=8)
        ws.row_dimensions[row + 1].height = 15 * (text.count("\n") + 1 + len(text) // 130)
        row += 3
    set_widths(ws, {"A": 2, "B": 24, "C": 16, "D": 16, "E": 16, "F": 16, "G": 16, "H": 16})


def build_inputs(wb: Workbook, results: dict) -> None:
    ws = wb.create_sheet("Inputs")
    ws["B2"] = "Inputs and fit"
    ws["B2"].font = TITLE_FONT
    labels = [
        (4, "Quote date (WSJ)", datetime.fromisoformat(results["quote_date"]), DATE_FMT, True),
        (5, "Market holiday (Labor Day)", datetime.combine(LABOR_DAY_2026, datetime.min.time()), DATE_FMT, True),
        (6, "Settlement date (T+1 business)", "=WORKDAY(B4,1,B5)", DATE_FMT, False),
        (7, "Face value", 100, "0", True),
        (8, "Coupons per year", 2, "0", True),
        (9, "Days per year (continuous time)", 365, "0", True),
        (10, "Min years to maturity in fit", results["svensson"]["min_years_in_fit"], "0.00", True),
    ]
    for row, label, value, fmt, is_input in labels:
        ws.cell(row=row, column=1, value=label)
        cell = ws.cell(row=row, column=2, value=value)
        cell.number_format = fmt
        if is_input:
            cell.fill = INPUT_FILL
        cell.border = BOX
    ws["A12"] = "Svensson parameters (continuous zero rate, decimals)"
    ws["A12"].font = BOLD
    params = results["svensson"]["params"]
    names = ["beta0", "beta1", "beta2", "beta3", "tau1", "tau2"]
    descriptions = {
        "beta0": "level: r(t) as t -> infinity",
        "beta1": "slope: r(0) = beta0 + beta1",
        "beta2": "first hump, decay tau1",
        "beta3": "second hump, decay tau2",
        "tau1": "decay (years) of slope/first hump",
        "tau2": "decay (years) of second hump",
    }
    for i, name in enumerate(names):
        row = 13 + i
        ws.cell(row=row, column=1, value=name)
        cell = ws.cell(row=row, column=2, value=float(params[name]))
        cell.number_format = "0.000000"
        cell.fill = INPUT_FILL
        cell.border = BOX
        ws.cell(row=row, column=3, value=descriptions[name])
        add_name(wb, NAME_OF[name], f"Inputs!$B${row}")
    add_name(wb, "Settle", "Inputs!$B$6")
    add_name(wb, "MinYears", "Inputs!$B$10")

    last = FIRST_ROW + results["svensson"]["n_bonds_total"] - 1
    stats = [
        (20, "Fit objective and statistics (bonds with In fit = 1)", None, None),
        (21, "Weighted SSE  = SUM( InFit x ((Pmodel - Pmkt) / ModDuration)^2 )", f"=SUM(Bonds!AA{FIRST_ROW}:AA{last})", "0.000000"),
        (22, "Bonds in fit", f"=SUM(Bonds!S{FIRST_ROW}:S{last})", "0"),
        (23, "Bonds total", f"=COUNT(Bonds!B{FIRST_ROW}:B{last})", "0"),
        (24, "Price RMSE (per 100 face)", f"=SQRT(SUMPRODUCT(Bonds!S{FIRST_ROW}:S{last},Bonds!V{FIRST_ROW}:V{last},Bonds!V{FIRST_ROW}:V{last})/B22)", "0.0000"),
        (25, "Yield RMSE (bp)", f"=SQRT(SUMPRODUCT(Bonds!S{FIRST_ROW}:S{last},Bonds!X{FIRST_ROW}:X{last},Bonds!X{FIRST_ROW}:X{last})/B22)", "0.00"),
        (26, "Yield MAE (bp)", f"=SUMPRODUCT(Bonds!S{FIRST_ROW}:S{last},ABS(Bonds!X{FIRST_ROW}:X{last}))/B22", "0.00"),
        (27, "Max |yield error| (bp)", f"=MAX(INDEX(Bonds!S{FIRST_ROW}:S{last}*ABS(Bonds!X{FIRST_ROW}:X{last}),0))", "0.00"),
        (28, "r(0) = beta0 + beta1 (%)", "=(Sv_beta0+Sv_beta1)*100", PCT3),
        (29, "r(30y) (%)", "=(" + ZERO_FORMULA.format(t="30") + ")*100", PCT3),
    ]
    for row, label, formula, fmt in stats:
        ws.cell(row=row, column=1, value=label).font = BOLD if formula is None else Font()
        if formula is not None:
            cell = ws.cell(row=row, column=2, value=formula)
            cell.number_format = fmt
            cell.border = BOX
    add_name(wb, "Weighted_SSE", "Inputs!$B$21")

    ws["A31"] = "Solver set-up"
    ws["A31"].font = BOLD
    ws["A32"] = ("Objective: $B$21 -> Min.  Changing cells: $B$13:$B$18.  Constraints: $B$13 >= 0, $B$17 >= 0.05, $B$18 >= 0.05.  "
                 "GRG Nonlinear; untick 'Make Unconstrained Variables Non-Negative'.")
    ws["A32"].alignment = WRAP
    ws.merge_cells("A32:E33")
    ws["A35"] = "Yellow cells are inputs; everything else is calculated."
    ws["A35"].fill = INPUT_FILL
    set_widths(ws, {"A": 62, "B": 18, "C": 34})


def build_data(wb: Workbook, quotes) -> None:
    ws = wb.create_sheet("Data")
    ws["A1"] = "U.S. Treasury Quotes - Treasury Notes & Bonds (WSJ, source Tullett Prebon). Prices in 32nds; third decimal = eighths of a 32nd."
    header(ws, 2, ["Maturity", "Coupon", "Bid", "Asked", "Chg", "Asked yield"])
    for i, q in enumerate(quotes.itertuples(index=False)):
        row = 3 + i
        ws.cell(row=row, column=1, value=datetime.combine(q.maturity, datetime.min.time())).number_format = DATE_FMT
        ws.cell(row=row, column=2, value=q.coupon).number_format = PCT3
        ws.cell(row=row, column=3, value=q.bid_quote).number_format = PCT3
        ws.cell(row=row, column=4, value=q.ask_quote).number_format = PCT3
        ws.cell(row=row, column=5, value=q.chg)
        ws.cell(row=row, column=6, value=q.asked_yield).number_format = PCT3
    ws.freeze_panes = "A3"
    set_widths(ws, {"A": 12, "B": 9, "C": 9, "D": 9, "E": 8, "F": 11})


BOND_COLUMNS = [
    ("#", 5), ("Maturity", 11), ("Coupon %", 8), ("Bid (WSJ)", 9), ("Ask (WSJ)", 9), ("WSJ asked yield %", 10),
    ("Ask clean price (decimal)", 12), ("Previous coupon", 11), ("Next coupon", 11), ("Days accrued (settle - prev)", 8),
    ("Days in period (next - prev)", 8), ("Accrued interest", 10), ("Dirty price", 11), ("Street yield % (Excel YIELD)", 11),
    ("Yield check vs WSJ (bp)", 9), ("Modified duration", 9), ("Remaining cash flows", 8), ("Years to maturity", 9),
    ("In fit (1/0)", 6), ("Model dirty price", 11), ("Model clean price", 11), ("Price error (model - mkt)", 10),
    ("Model yield %", 10), ("Yield error (bp)", 9), ("Weight = 1 / duration", 9), ("Weighted residual", 10),
    ("Weighted sq. residual (in fit)", 11),
]


def build_bonds(wb: Workbook, n_bonds: int, n_flow_cols: int) -> None:
    ws = wb.create_sheet("Bonds")
    ws["A1"] = ("Per-bond mechanics with Excel bond functions (basis 1 = actual/actual, frequency 2). Coupon-period days are taken as "
                "COUPNCD - COUPPCD because COUPDAYS(basis 1) mis-states some month-end periods (183 vs 184 days). Model prices come from the CashFlows sheet.")
    header(ws, 3, [c[0] for c in BOND_COLUMNS])
    ws.row_dimensions[3].height = 45
    pv_first = get_column_letter(6 + 2 * n_flow_cols + 2)  # PV grid start on CashFlows
    pv_last = get_column_letter(6 + 3 * n_flow_cols + 1)
    for i in range(n_bonds):
        r = FIRST_ROW + i
        d = 3 + i  # Data row
        t = f"R{r}"
        formulas = {
            "A": i + 1,
            "B": f"=Data!A{d}",
            "C": f"=Data!B{d}",
            "D": f"=Data!C{d}",
            "E": f"=Data!D{d}",
            "F": f"=Data!F{d}",
            "G": f'=INT(E{r})+(VALUE(MID(TEXT(E{r},"0.000"),FIND(".",TEXT(E{r},"0.000"))+1,2))+VALUE(RIGHT(TEXT(E{r},"0.000"),1))/8)/32',
            "H": f"=COUPPCD(Settle,B{r},2,1)",
            "I": f"=COUPNCD(Settle,B{r},2,1)",
            "J": f"=Settle-H{r}",
            "K": f"=I{r}-H{r}",
            "L": f"=C{r}/2*J{r}/K{r}",
            "M": f"=G{r}+L{r}",
            "N": f"=YIELD(Settle,B{r},C{r}/100,G{r},100,2,1)*100",
            "O": f"=(N{r}-F{r})*100",
            "P": f"=MDURATION(Settle,B{r},C{r}/100,N{r}/100,2,1)",
            "Q": f"=COUPNUM(Settle,B{r},2,1)",
            "R": f"=(B{r}-Settle)/365",
            "S": f"=IF({t}>=MinYears,1,0)",
            "T": f"=SUM(CashFlows!{pv_first}{r}:{pv_last}{r})",
            "U": f"=T{r}-L{r}",
            "V": f"=U{r}-G{r}",
            "W": f"=YIELD(Settle,B{r},C{r}/100,U{r},100,2,1)*100",
            "X": f"=(W{r}-N{r})*100",
            "Y": f"=1/P{r}",
            "Z": f"=Y{r}*(T{r}-M{r})",
            "AA": f"=S{r}*Z{r}^2",
        }
        formats = {"B": DATE_FMT, "C": PCT3, "D": PCT3, "E": PCT3, "F": PCT3, "G": PX, "H": DATE_FMT, "I": DATE_FMT,
                   "L": PX, "M": PX, "N": PCT3, "O": "0.00", "P": "0.0000", "R": "0.0000", "T": PX, "U": PX,
                   "V": "0.0000", "W": PCT3, "X": "0.00", "Y": "0.0000", "Z": "0.000000", "AA": "0.00000000"}
        for col, value in formulas.items():
            cell = ws[f"{col}{r}"]
            cell.value = value
            if col in formats:
                cell.number_format = formats[col]
    ws.freeze_panes = f"C{FIRST_ROW}"
    set_widths(ws, {get_column_letter(i + 1): w for i, (_, w) in enumerate(BOND_COLUMNS)})


def build_cashflows(wb: Workbook, n_bonds: int, n_flow_cols: int) -> None:
    ws = wb.create_sheet("CashFlows")
    ws["A1"] = ("Every remaining payment of every bond. Block 1: payment dates (k-th coupon after settlement). "
                "Block 2: Svensson zero rate r(t_k). Block 3: present value = amount x exp(-r t). Model dirty price = row sum of block 3.")
    base = ["#", "Maturity", "Coupon %", "Remaining flows N", "Month-end (1/0)"]
    header(ws, 3, base)
    date_start = 6
    rate_start = date_start + n_flow_cols + 1
    pv_start = rate_start + n_flow_cols + 1
    header(ws, 3, [f"Date {k}" for k in range(1, n_flow_cols + 1)], date_start)
    header(ws, 3, [f"r(t) {k}" for k in range(1, n_flow_cols + 1)], rate_start)
    header(ws, 3, [f"PV {k}" for k in range(1, n_flow_cols + 1)], pv_start)
    ws.cell(row=2, column=date_start, value="Payment dates").font = BOLD
    ws.cell(row=2, column=rate_start, value="Zero rates (continuous, decimal)").font = BOLD
    ws.cell(row=2, column=pv_start, value="Present values (per 100 face)").font = BOLD

    for i in range(n_bonds):
        r = FIRST_ROW + i
        ws[f"A{r}"] = i + 1
        ws[f"B{r}"] = f"=Bonds!B{r}"
        ws[f"B{r}"].number_format = DATE_FMT
        ws[f"C{r}"] = f"=Bonds!C{r}"
        ws[f"C{r}"].number_format = PCT3
        ws[f"D{r}"] = f"=Bonds!Q{r}"
        ws[f"E{r}"] = f"=IF(DAY(B{r})=DAY(EOMONTH(B{r},0)),1,0)"
        for k in range(1, n_flow_cols + 1):
            dcol = get_column_letter(date_start + k - 1)
            rcol = get_column_letter(rate_start + k - 1)
            pcol = get_column_letter(pv_start + k - 1)
            shifted = f"EDATE(Bonds!$I{r},6*({k}-1))"
            ws[f"{dcol}{r}"] = f'=IF({k}>$D{r},"",IF($E{r}=1,EOMONTH({shifted},0),{shifted}))'
            ws[f"{dcol}{r}"].number_format = DATE_FMT
            t = f"(({dcol}{r}-Settle)/365)"
            ws[f"{rcol}{r}"] = f'=IF({dcol}{r}="","",{ZERO_FORMULA.format(t=t)})'
            ws[f"{rcol}{r}"].number_format = "0.00000"
            amount = f"($C{r}/2+IF({k}=$D{r},100,0))"
            ws[f"{pcol}{r}"] = f'=IF({dcol}{r}="",0,{amount}*EXP(-{rcol}{r}*{t}))'
            ws[f"{pcol}{r}"].number_format = "0.0000"
    ws.freeze_panes = f"F{FIRST_ROW}"
    set_widths(ws, {"A": 5, "B": 11, "C": 8, "D": 7, "E": 7})
    for c in range(date_start, pv_start + n_flow_cols):
        ws.column_dimensions[get_column_letter(c)].width = 10.5


def build_zero_curve(wb: Workbook, results: dict, n_bonds: int) -> None:
    ws = wb.create_sheet("ZeroCurve")
    ws["A1"] = ("Deliverable: continuously compounded discount rate for every Treasury payment date in the sample "
                "(union of all coupon and principal dates). Dates are the distinct payment dates from the CashFlows sheet; rates are live formulas.")
    header(ws, 3, ["Payment date", "Days from settlement", "t (years)", "Zero rate r(t) %", "Discount factor d(t)", "Instantaneous forward f(t) %"])
    ws.row_dimensions[3].height = 32
    payments = results["payment_dates"]
    for i, p in enumerate(payments):
        r = 4 + i
        ws.cell(row=r, column=1, value=datetime.fromisoformat(p["payment_date"])).number_format = DATE_FMT
        ws.cell(row=r, column=2, value=f"=A{r}-Settle").number_format = "0"
        ws.cell(row=r, column=3, value=f"=B{r}/365").number_format = "0.0000"
        ws.cell(row=r, column=4, value="=(" + ZERO_FORMULA.format(t=f"C{r}") + ")*100").number_format = "0.0000"
        ws.cell(row=r, column=5, value=f"=EXP(-D{r}/100*C{r})").number_format = "0.000000"
        ws.cell(row=r, column=6, value="=(" + FORWARD_FORMULA.format(t=f"C{r}") + ")*100").number_format = "0.0000"
    last = 3 + len(payments)
    ws.freeze_panes = "A4"
    set_widths(ws, {"A": 13, "B": 11, "C": 10, "D": 13, "E": 14, "F": 16})

    chart = ScatterChart()
    chart.title = "ZCB term structure: zero rate, instantaneous forward and Treasury street yields"
    chart.style = 13
    chart.x_axis.title = "Years from settlement"
    chart.y_axis.title = "Rate (% per year)"
    chart.x_axis.scaling.min = 0
    chart.x_axis.scaling.max = 31
    chart.y_axis.scaling.min = 2.5
    chart.y_axis.scaling.max = 6.0
    chart.height, chart.width = 11, 22
    x_ref = Reference(ws, min_col=3, min_row=4, max_row=last)
    zero = Series(Reference(ws, min_col=4, min_row=4, max_row=last), x_ref, title="Zero rate r(t)")
    zero.marker.symbol = "none"
    zero.graphicalProperties.line.solidFill = "2A78D6"
    zero.graphicalProperties.line.width = 22000
    fwd = Series(Reference(ws, min_col=6, min_row=4, max_row=last), x_ref, title="Instantaneous forward f(t)")
    fwd.marker.symbol = "none"
    fwd.graphicalProperties.line.solidFill = "EB6834"
    fwd.graphicalProperties.line.dashStyle = "dash"
    fwd.graphicalProperties.line.width = 22000
    bonds_ws = wb["Bonds"]
    b_last = FIRST_ROW + n_bonds - 1
    ytm = Series(Reference(bonds_ws, min_col=14, min_row=FIRST_ROW, max_row=b_last),
                 Reference(bonds_ws, min_col=18, min_row=FIRST_ROW, max_row=b_last), title="Treasury street yields")
    ytm.marker.symbol = "circle"
    ytm.marker.size = 4
    ytm.marker.graphicalProperties.solidFill = "898781"
    ytm.marker.graphicalProperties.line.solidFill = "898781"
    ytm.graphicalProperties.line.noFill = True
    for s in (zero, fwd, ytm):
        chart.series.append(s)
    ws.add_chart(chart, "H3")


def build_robustness(wb: Workbook, results: dict) -> None:
    ws = wb.create_sheet("Robustness")
    ws["A1"] = "Robustness checks (values computed by the Python implementation, python/run_analysis.py)."
    ws["A1"].font = BOLD
    row = 3
    ws.cell(row=row, column=1, value="1. Convention check: RMSE of our recomputed asked yield vs the WSJ 'Asked yield' column").font = BOLD
    header(ws, row + 1, ["Settlement rule", "Settlement", "Price format", "RMSE (bp)", "Median |err| (bp)", "Max |err| (bp)"])
    for i, c in enumerate(results["conventions"]):
        r = row + 2 + i
        ws.cell(row=r, column=1, value=c["settlement_rule"])
        ws.cell(row=r, column=2, value=datetime.fromisoformat(c["settlement"])).number_format = DATE_FMT
        ws.cell(row=r, column=3, value=c["price_format"])
        ws.cell(row=r, column=4, value=c["rmse_bp"]).number_format = "0.000"
        ws.cell(row=r, column=5, value=c["median_abs_bp"]).number_format = "0.000"
        ws.cell(row=r, column=6, value=c["max_abs_bp"]).number_format = "0.000"
    row = row + 3 + len(results["conventions"])
    ws.cell(row=row, column=1, value="2. Model comparison (bonds in fit)").font = BOLD
    header(ws, row + 1, ["Model", "Bonds in fit", "Weighted SSE", "Price RMSE", "Yield RMSE (bp)", "Yield MAE (bp)", "Max |yield err| (bp)"])
    for i, key in enumerate(("svensson", "nelson_siegel")):
        s = results[key]
        r = row + 2 + i
        ws.cell(row=r, column=1, value="Svensson (6 parameters)" if key == "svensson" else "Nelson-Siegel (4 parameters)")
        ws.cell(row=r, column=2, value=s["n_bonds_in_fit"])
        ws.cell(row=r, column=3, value=s["weighted_sse"]).number_format = "0.0000"
        ws.cell(row=r, column=4, value=s["price_rmse"]).number_format = "0.0000"
        ws.cell(row=r, column=5, value=s["yield_rmse_bp"]).number_format = "0.00"
        ws.cell(row=r, column=6, value=s["yield_mae_bp"]).number_format = "0.00"
        ws.cell(row=r, column=7, value=s["yield_max_abs_bp"]).number_format = "0.00"
    row += 5
    ws.cell(row=row, column=1, value="3. beta0 profile: refit with the asymptotic level pinned; the in-sample curve barely moves").font = BOLD
    cols = list(results["beta0_profile"][0].keys())
    header(ws, row + 1, cols)
    for i, p in enumerate(results["beta0_profile"]):
        r = row + 2 + i
        for j, c in enumerate(cols):
            ws.cell(row=r, column=1 + j, value=p[c]).number_format = "0.0000"
    set_widths(ws, {"A": 30, "B": 14, "C": 16, "D": 14, "E": 16, "F": 16, "G": 18})
    for col in "HIJKLMN":
        ws.column_dimensions[col].width = 13


def build_methodology(wb: Workbook, text: str) -> None:
    ws = wb.create_sheet("Methodology")
    ws["B2"] = "Methodology (one page)"
    ws["B2"].font = TITLE_FONT
    for i, paragraph in enumerate(text.strip().split("\n\n")):
        cell = ws.cell(row=4 + i, column=2, value=paragraph)
        cell.alignment = WRAP
        ws.merge_cells(start_row=4 + i, start_column=2, end_row=4 + i, end_column=9)
        ws.row_dimensions[4 + i].height = max(18, 15 * (len(paragraph) // 120 + 1))
    set_widths(ws, {"A": 2, "B": 18, "C": 18, "D": 18, "E": 18, "F": 18, "G": 18, "H": 18, "I": 18})


def build_workbook(data_path: Path, results_path: Path, methodology_path: Path, out_path: Path) -> Path:
    results = json.loads(results_path.read_text())
    sheet = load_wsj_quotes(data_path)
    n_bonds = len(sheet.quotes)
    n_flow_cols = max(b["n_cash_flows"] for b in results["bonds"])
    wb = Workbook()
    build_cover(wb, results)
    build_inputs(wb, results)
    build_data(wb, sheet.quotes)
    build_bonds(wb, n_bonds, n_flow_cols)
    build_cashflows(wb, n_bonds, n_flow_cols)
    build_zero_curve(wb, results, n_bonds)
    build_robustness(wb, results)
    build_methodology(wb, methodology_path.read_text() if methodology_path.exists() else "See README.")
    wb.calculation.fullCalcOnLoad = True
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    return out_path


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--data", default="../data/Treasury_data_090426.xlsx")
    parser.add_argument("--results", default="../output/results.json")
    parser.add_argument("--methodology", default="../writeup/methodology_excel.txt")
    parser.add_argument("--out", default="../excel/Miniproject2_ZCB_Term_Structure.xlsx")
    args = parser.parse_args(argv)
    path = build_workbook(Path(args.data), Path(args.results), Path(args.methodology), Path(args.out))
    print(f"Wrote {path.resolve()} ({path.stat().st_size / 1e6:.1f} MB)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
