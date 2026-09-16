#!/usr/bin/env python3
"""Generate the one-page methodology write-up from results.json so every
number in the text matches the analysis. Produces Markdown (repo/site), plain
text (Excel 'Methodology' sheet) and print-ready HTML (for the PDF)."""
from __future__ import annotations

import argparse
import html
import json
import math
import shutil
import subprocess
from datetime import date
from pathlib import Path

TITLE = "Miniproject 2 - Creating a ZCB Term Structure"
SUBTITLE = "FRE 6103 Valuation for Financial Engineering, NYU Tandon | Raj Pawar | Methodology"


def _fmt_date(iso: str) -> str:
    return date.fromisoformat(iso).strftime("%b %d, %Y").replace(" 0", " ")


def _zero_at(results: dict, years: float) -> float:
    rows = results["payment_dates"]
    return min(rows, key=lambda r: abs(r["t_years"] - years))["zero_rate_cc_pct"]


def build_paragraphs(results: dict) -> list[tuple[str, str]]:
    sv, ns = results["svensson"], results["nelson_siegel"]
    p = sv["params"]
    conv = results["conventions"]
    best = conv[0]
    next_best = min(
        (c for c in conv if c["settlement"] != best["settlement"] and c["price_format"] == best["price_format"]),
        key=lambda c: c["rmse_bp"],
    )
    n_total, n_fit = sv["n_bonds_total"], sv["n_bonds_in_fit"]
    prof = {round(r["beta0_fixed_pct"]): r for r in results["beta0_profile"]}
    horizons = [k for k in prof[0] if k.startswith("zero_")]
    max_shift_bp = max(abs(prof[3][h] - prof[0][h]) for h in horizons) * 100
    d30_change = abs(math.exp(-prof[3]["zero_30y_pct"] / 100 * 30) / math.exp(-prof[0]["zero_30y_pct"] / 100 * 30) - 1) * 100
    rmse_shift = prof[3]["approx_yield_rmse_bp"] - prof[0]["approx_yield_rmse_bp"]
    bonds = results["bonds"]
    worst = max((b for b in bonds if b["in_fit"]), key=lambda b: abs(b["yield_error_bp"]))
    pay = results["payment_dates"]
    r20 = _zero_at(results, 20.0)
    hump = max(pay, key=lambda r: r["zero_rate_cc_pct"])

    data = (
        f"Wall Street Journal / Tullett Prebon 3 pm quotes for all {n_total} US Treasury notes and bonds on "
        f"{_fmt_date(results['quote_date'])} (the Week 2 data file). Quotes are clean prices in 32nds with a third "
        f"decimal in eighths of a 32nd (99.256 = 99 + 25.75/32); we price off the asked side. Regular-way settlement "
        f"is T+1 business days, {_fmt_date(results['settlement'])} (Monday Sep 7 is Labor Day). Both conventions were "
        f"verified by recomputing the WSJ asked yield of every bond: median discrepancy {best['median_abs_bp']:.2f} bp, "
        f"RMSE {best['rmse_bp']:.1f} bp, versus {next_best['rmse_bp']:.1f} bp for the next-best settlement rule and over "
        f"{min(c['rmse_bp'] for c in conv if c['price_format'] == 'plain decimal'):.0f} bp if the quotes are read as decimals."
    )
    cash_flows = (
        "Each bond pays coupon/2 semi-annually on its maturity day-of-month (month-end maturities pay on month-ends) "
        "and 100 at maturity. Accrued interest is actual/actual within the coupon period, coupon/2 x days accrued / "
        "days in period, and dirty price = clean price + accrued (slide 34). Time to each payment is t = days/365 from "
        "settlement. Street yields use the Treasury convention (semi-annual compounding with a fractional first period, "
        "simple interest when one payment remains), which reproduces Excel's YIELD."
    )
    model = (
        "The continuously compounded zero rate follows the Svensson (1994) form suggested in the data file, "
        "r(t) = b0 + b1 (1 - e^(-t/T1))/(t/T1) + b2 [(1 - e^(-t/T1))/(t/T1) - e^(-t/T1)] + b3 [(1 - e^(-t/T2))/(t/T2) - e^(-t/T2)]. "
        "The discount factor is d(t) = e^(-r(t) t) and each bond's model dirty price is the sumproduct of its payment "
        "stream with the discount factors, sum_k CF_k d(t_k) (slide 33). The form nests Nelson-Siegel (b3 = 0), is smooth, "
        "has finite limits at both ends (r(0) = b0 + b1, r(inf) = b0) and gives the instantaneous forward curve in closed "
        "form, f(t) = b0 + b1 e^(-t/T1) + b2 (t/T1) e^(-t/T1) + b3 (t/T2) e^(-t/T2)."
    )
    estimation = (
        "The six parameters minimise sum_i [(P_i(model) - P_i(market)) / (P_i(market) D_i)]^2, each price error divided by "
        "dirty price times modified duration, which is the first-order yield error, so the objective is the sum of squared "
        "yield errors (slide 31) while each model price stays a linear sumproduct. Bonds with under {mn:.0f} months to maturity "
        "({nx} of {nt}) are excluded from estimation, as in Gurkaynak, Sack and Wright (2007): their prices reflect "
        "money-market conditions and, with near-zero duration, a one-cent quote error is a 50 bp yield error. They are "
        "still priced off the curve as an out-of-sample check. Constraints: b0 >= 0, T1, T2 >= 0.05. Because the "
        "objective is multimodal in (T1, T2), the least-squares solver is run from {ns} starting points on a tau grid; "
        "{nopt} of them reach the same optimum. In Excel the identical objective is the sum-of-squares cell on the Curve sheet, "
        "minimised with Solver (GRG Nonlinear); the workbook reproduces every Python model price, yield and zero rate."
    ).format(mn=sv["min_years_in_fit"] * 12, nx=n_total - n_fit, nt=n_total, ns=sv["n_starts"], nopt=sv["n_starts_at_optimum"])
    results_text = (
        f"b0 = {p['beta0']:.4f} (non-negativity bound active), b1 = {p['beta1']:.4f}, b2 = {p['beta2']:.4f}, "
        f"b3 = {p['beta3']:.4f}, T1 = {p['tau1']:.2f}, T2 = {p['tau2']:.2f}. The curve prices the {n_fit} bonds with a "
        f"yield RMSE of {sv['yield_rmse_bp']:.1f} bp (mean absolute {sv['yield_mae_bp']:.1f} bp; worst {abs(worst['yield_error_bp']):.1f} bp "
        f"for the {worst['label']} bond). Nelson-Siegel on the same data gives {ns['yield_rmse_bp']:.1f} bp, so the second "
        f"hump is needed. The zero curve rises from {(p['beta0'] + p['beta1']) * 100:.2f}% at the short end to "
        f"{_zero_at(results, 1):.2f}% at 1 year, {_zero_at(results, 5):.2f}% at 5, {_zero_at(results, 10):.2f}% at 10, "
        f"{r20:.2f}% at 20 and {_zero_at(results, 30):.2f}% at 30 years, peaking at {hump['zero_rate_cc_pct']:.2f}% around "
        f"{hump['t_years']:.0f} years. The deliverable table lists r(t), d(t) and f(t) for all {len(pay)} distinct Treasury "
        f"payment dates from {_fmt_date(pay[0]['payment_date'])} to {_fmt_date(pay[-1]['payment_date'])}. The asymptotic "
        f"level b0 is weakly identified by a 30-year sample: pinning it anywhere from 0% to 3% worsens the fit by only "
        f"{rmse_shift:.1f} bp RMSE and moves the in-sample curve by about {max_shift_bp:.0f} bp at the two ends (a "
        f"{d30_change:.0f}% change in d(30y)), so the payment-date discount factors are robust to it; with b0 = 0 the fitted "
        f"forward rate decays beyond the sample, so the curve should not be extrapolated past 30 years."
    )
    return [
        ("Data and conventions", data),
        ("Cash flows", cash_flows),
        ("Model", model),
        ("Estimation", estimation),
        ("Results", results_text),
    ]


def to_markdown(paragraphs) -> str:
    out = [f"# {TITLE}", "", f"_{SUBTITLE}_", ""]
    for heading, text in paragraphs:
        out += [f"**{heading}.** {text}", ""]
    return "\n".join(out)


def to_plain(paragraphs) -> str:
    return "\n\n".join(f"{heading}. {text}" for heading, text in paragraphs)


def to_html(paragraphs) -> str:
    body = "\n".join(
        f"<p><strong>{html.escape(h)}.</strong> {html.escape(t)}</p>" for h, t in paragraphs
    )
    return f"""<!doctype html><html><head><meta charset="utf-8"><title>{html.escape(TITLE)}</title>
<style>
@page {{ size: Letter; margin: 0.7in 0.8in; }}
body {{ font-family: -apple-system, "Helvetica Neue", Helvetica, Arial, sans-serif; color: #0b0b0b; font-size: 10.6pt; line-height: 1.38; margin: 0; }}
h1 {{ font-size: 17pt; margin: 0 0 2pt 0; color: #1f3a5f; }}
.sub {{ color: #52514e; font-size: 9.5pt; margin: 0 0 12pt 0; }}
p {{ margin: 0 0 8pt 0; text-align: justify; }}
strong {{ color: #1f3a5f; }}
</style></head><body>
<h1>{html.escape(TITLE)}</h1>
<p class="sub">{html.escape(SUBTITLE)}</p>
{body}
</body></html>"""


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--results", default="../output/results.json")
    parser.add_argument("--out-dir", default="../writeup")
    parser.add_argument("--no-pdf", action="store_true", help="skip the headless-Chrome PDF step")
    args = parser.parse_args(argv)
    results = json.loads(Path(args.results).read_text())
    paragraphs = build_paragraphs(results)
    out = Path(args.out_dir)
    out.mkdir(parents=True, exist_ok=True)
    (out / "Miniproject2_ZCB_Term_Structure_Writeup.md").write_text(to_markdown(paragraphs))
    (out / "methodology_excel.txt").write_text(to_plain(paragraphs))
    html_path = out / "Miniproject2_ZCB_Term_Structure_Writeup.html"
    html_path.write_text(to_html(paragraphs))
    words = sum(len(t.split()) for _, t in paragraphs)
    print(f"Methodology written to {out.resolve()} ({words} words)")
    if not args.no_pdf:
        print_pdf(html_path, out / "Miniproject2_ZCB_Term_Structure_Writeup.pdf")
    return 0


CHROME_CANDIDATES = (
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    "google-chrome",
    "chromium",
    "chrome",
)


def print_pdf(html_path: Path, pdf_path: Path) -> None:
    """Print the HTML write-up to a one-page PDF with headless Chrome, if available."""
    chrome = next((c for c in CHROME_CANDIDATES if Path(c).exists() or shutil.which(c)), None)
    if chrome is None:
        print("Chrome not found: PDF not regenerated (open the HTML and print to PDF instead)")
        return
    subprocess.run(
        [chrome, "--headless=new", "--disable-gpu", "--no-pdf-header-footer",
         f"--print-to-pdf={pdf_path.resolve()}", html_path.resolve().as_uri()],
        check=True, capture_output=True,
    )
    print(f"PDF written to {pdf_path.resolve()}")


if __name__ == "__main__":
    raise SystemExit(main())
