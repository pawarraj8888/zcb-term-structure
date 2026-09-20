#!/usr/bin/env python3
"""Generate the write-up from output/results.json so every number matches the
analysis. Produces Markdown (repo/site), a print-ready HTML document and a PDF
(headless Chrome): page 1 methodology (under one page), page 2 figure and
tables, appendix with the discount rate at every Treasury payment date."""
from __future__ import annotations

import argparse
import html as html_lib
import json
import math
import re
import shutil
import subprocess
import sys
from datetime import date
from pathlib import Path

import numpy as np

sys.path.insert(0, str(Path(__file__).resolve().parent))
from zcb.curves import discount_factor, svensson_forward, svensson_zero  # noqa: E402

TITLE = "Miniproject 2 - Creating a ZCB Term Structure"
COURSE = "FRE 6103 Valuation for Financial Engineering, NYU Tandon"
AUTHOR = "Monalisa Maity (mm16178) and Raj Pawar (rsp9234)"
FIGURE = "../output/figures/zero_curve.png"
SELECTED_MATURITIES = (0.25, 0.5, 1, 2, 3, 5, 7, 10, 15, 20, 25, 30)
KATEX = "https://cdn.jsdelivr.net/npm/katex@0.16.11/dist"

SVENSSON_TEX = (
    r"r(t)=\beta_0+\beta_1\,\frac{1-e^{-t/\tau_1}}{t/\tau_1}"
    r"+\beta_2\left(\frac{1-e^{-t/\tau_1}}{t/\tau_1}-e^{-t/\tau_1}\right)"
    r"+\beta_3\left(\frac{1-e^{-t/\tau_2}}{t/\tau_2}-e^{-t/\tau_2}\right)"
    r"\\[2pt] d(t)=e^{-r(t)\,t},\qquad P_i^{\mathrm{model}}=\sum_k CF_{ik}\,d(t_{ik})"
)
OBJECTIVE_TEX = (
    r"\min_{\beta,\tau}\;\sum_{i\in\text{fit}}\left(\frac{P_i^{\mathrm{model}}-P_i^{\mathrm{market}}}"
    r"{P_i^{\mathrm{market}}\,D_i}\right)^2,\qquad \beta_0\ge 0,\;\tau_1,\tau_2\ge 0.05"
)
FORWARD_TEX = (
    r"f(t)=\beta_0+\beta_1 e^{-t/\tau_1}+\beta_2\frac{t}{\tau_1}e^{-t/\tau_1}+\beta_3\frac{t}{\tau_2}e^{-t/\tau_2}"
)


def _fmt_date(iso: str, long: bool = False) -> str:
    d = date.fromisoformat(iso)
    return d.strftime("%A, %B %d, %Y" if long else "%b %d, %Y").replace(" 0", " ")


def _zero_at(results: dict, years: float) -> float:
    return min(results["payment_dates"], key=lambda r: abs(r["t_years"] - years))["zero_rate_cc_pct"]


def build_sections(results: dict) -> list[dict]:
    """Methodology paragraphs as HTML (inline math with Unicode/sup/sub) plus optional display equations."""
    sv, ns = results["svensson"], results["nelson_siegel"]
    p = sv["params"]
    conv = results["conventions"]
    best = conv[0]
    next_best = min((c for c in conv if c["settlement"] != best["settlement"] and c["price_format"] == best["price_format"]),
                    key=lambda c: c["rmse_bp"])
    decimal = min(c["rmse_bp"] for c in conv if c["price_format"] == "plain decimal")
    n_total, n_fit = sv["n_bonds_total"], sv["n_bonds_in_fit"]
    prof = {round(r["beta0_fixed_pct"]): r for r in results["beta0_profile"]}
    horizons = [k for k in prof[0] if k.startswith("zero_")]
    max_shift_bp = max(abs(prof[3][h] - prof[0][h]) for h in horizons) * 100
    d30_change = abs(math.exp(-prof[3]["zero_30y_pct"] / 100 * 30) / math.exp(-prof[0]["zero_30y_pct"] / 100 * 30) - 1) * 100
    rmse_shift = prof[3]["approx_yield_rmse_bp"] - prof[0]["approx_yield_rmse_bp"]
    worst = max((b for b in results["bonds"] if b["in_fit"]), key=lambda b: abs(b["yield_error_bp"]))
    pay = results["payment_dates"]
    hump = max(pay, key=lambda r: r["zero_rate_cc_pct"])

    return [
        {"heading": "Data and conventions", "html": (
            f"Wall Street Journal (Tullett Prebon) 3 pm quotes for all {n_total} US Treasury notes and bonds on "
            f"{_fmt_date(results['quote_date'], long=True)}, the Week 2 data file. Quotes are clean prices in 32nds with a "
            f"third decimal in eighths of a 32nd (99.256 = 99 + 25.75/32); we price off the asked side. Regular-way settlement "
            f"is T+1 business days, {_fmt_date(results['settlement'], long=True)} (Monday is Labor Day). Both conventions were "
            f"checked by recomputing every WSJ asked yield: median discrepancy {best['median_abs_bp']:.2f} bp and RMSE "
            f"{best['rmse_bp']:.1f} bp, against {next_best['rmse_bp']:.1f} bp for the next-best settlement date and over "
            f"{decimal:.0f} bp if the quotes are read as decimals.")},
        {"heading": "Cash flows", "html": (
            "Each bond pays coupon/2 semi-annually on its maturity day-of-month (month-end maturities pay on month-ends) and "
            "100 at maturity. Accrued interest is actual/actual within the coupon period, coupon/2 × days accrued / days in "
            "period; dirty price = clean price + accrued (slide 34). Time to each payment is <i>t</i> = days/365 from settlement. "
            "Street yields follow the Treasury convention (semi-annual compounding with a fractional first period, simple "
            "interest when one payment remains), which reproduces Excel's YIELD.")},
        {"heading": "Model", "html": (
            "The continuously compounded zero rate follows the Svensson (1994) form suggested with the data file, the discount "
            "factor is <i>d</i>(<i>t</i>) = <i>e</i><sup>−<i>r</i>(<i>t</i>)<i>t</i></sup>, and each bond's model dirty price is the "
            "sumproduct of its payment stream with the discount factors (slide 33):"),
            "display": SVENSSON_TEX, "after": (
            "The form nests Nelson-Siegel (β<sub>3</sub> = 0), is smooth, has finite limits at both ends "
            "(<i>r</i>(0) = β<sub>0</sub> + β<sub>1</sub>, <i>r</i>(∞) = β<sub>0</sub>) and gives the instantaneous forward "
            "curve in closed form, <i>f</i>(<i>t</i>) = β<sub>0</sub> + β<sub>1</sub><i>e</i><sup>−<i>t</i>/τ<sub>1</sub></sup> + "
            "β<sub>2</sub>(<i>t</i>/τ<sub>1</sub>)<i>e</i><sup>−<i>t</i>/τ<sub>1</sub></sup> + "
            "β<sub>3</sub>(<i>t</i>/τ<sub>2</sub>)<i>e</i><sup>−<i>t</i>/τ<sub>2</sub></sup>.")},
        {"heading": "Estimation", "html": (
            "The six parameters minimise the sum of squared price errors, each divided by dirty price times modified "
            "duration <i>D<sub>i</sub></i>, which is the first-order yield error; the objective is therefore the sum of squared "
            "yield errors (slide 31) while every model price stays a linear sumproduct:"),
            "display": OBJECTIVE_TEX, "after": (
            f"Bonds with under 3 months to maturity ({n_total - n_fit} of {n_total}) are excluded from estimation, following "
            f"Gürkaynak, Sack and Wright (2007): their prices reflect money-market conditions and, with near-zero duration, a "
            f"one-cent quote error is a 50 bp yield error; they are still priced off the curve as an out-of-sample check. The "
            f"objective is multimodal in (τ<sub>1</sub>, τ<sub>2</sub>), so the least-squares solver is started from "
            f"{sv['n_starts']} points on a τ grid; {sv['n_starts_at_optimum']} reach the same optimum. In Excel the same "
            f"objective is the sum-of-squares cell on the Curve sheet, minimised with Solver (GRG Nonlinear); the workbook "
            f"reproduces every Python model price, yield and zero rate.")},
        {"heading": "Results", "html": (
            f"β<sub>0</sub> = {p['beta0']:.4f} (bound active), β<sub>1</sub> = {p['beta1']:.4f}, β<sub>2</sub> = {p['beta2']:.4f}, "
            f"β<sub>3</sub> = {p['beta3']:.4f}, τ<sub>1</sub> = {p['tau1']:.2f}, τ<sub>2</sub> = {p['tau2']:.2f} (Table 1). The curve "
            f"prices the {n_fit} bonds with a yield RMSE of {sv['yield_rmse_bp']:.1f} bp (mean absolute {sv['yield_mae_bp']:.1f} bp; "
            f"worst {abs(worst['yield_error_bp']):.1f} bp, the {worst['coupon']:.3f}% {_fmt_date(worst['maturity'])} note); "
            f"Nelson-Siegel gives {ns['yield_rmse_bp']:.1f} bp, so the second hump is needed. The zero curve rises from "
            f"{(p['beta0'] + p['beta1']) * 100:.2f}% at the short end to {_zero_at(results, 1):.2f}% at 1 year, "
            f"{_zero_at(results, 5):.2f}% at 5, {_zero_at(results, 10):.2f}% at 10, {_zero_at(results, 20):.2f}% at 20 and "
            f"{_zero_at(results, 30):.2f}% at 30 years, peaking at {hump['zero_rate_cc_pct']:.2f}% near {hump['t_years']:.0f} years "
            f"(Figure 1, Table 2). Table 3 lists <i>r</i>(<i>t</i>), <i>d</i>(<i>t</i>) and <i>f</i>(<i>t</i>) for all {len(pay)} "
            f"distinct Treasury payment dates from {_fmt_date(pay[0]['payment_date'])} to {_fmt_date(pay[-1]['payment_date'])}. "
            f"β<sub>0</sub>, the rate at infinite maturity, is weakly identified by a 30-year sample: pinning it anywhere from 0% "
            f"to 3% worsens the fit by only {rmse_shift:.1f} bp RMSE and moves the in-sample curve by about {max_shift_bp:.0f} bp "
            f"at the two ends ({d30_change:.0f}% in <i>d</i>(30y)), so the payment-date discount factors are robust to it; with "
            f"β<sub>0</sub> = 0 the forward rate decays beyond the sample, so the curve should not be extrapolated past 30 years.")},
    ]


def _strip(html: str) -> str:
    text = re.sub(r"<sup>(.*?)</sup>", r"^\1", html)
    text = re.sub(r"<sub>(.*?)</sub>", r"_\1", text)
    return html_lib.unescape(re.sub(r"<[^>]+>", "", text))


def build_paragraphs(results: dict) -> list[tuple[str, str]]:
    """Plain-text (heading, text) pairs, kept for callers that need prose only."""
    out = []
    for s in build_sections(results):
        text = _strip(s["html"])
        if s.get("after"):
            text += " " + _strip(s["after"])
        out.append((s["heading"], text))
    return out


# ------------------------------------------------------------------ tables
def parameter_rows(results: dict) -> list[tuple[str, str, str]]:
    sv, ns = results["svensson"], results["nelson_siegel"]
    ps, pn = sv["params"], ns["params"]
    rows = [(f"β<sub>{i}</sub>", f"{ps[f'beta{i}']:.5f}", f"{pn[f'beta{i}']:.5f}") for i in range(4)]
    rows += [("τ<sub>1</sub> (years)", f"{ps['tau1']:.3f}", f"{pn['tau1']:.3f}"),
             ("τ<sub>2</sub> (years)", f"{ps['tau2']:.3f}", "—")]
    rows += [
        ("Bonds in fit", str(sv["n_bonds_in_fit"]), str(ns["n_bonds_in_fit"])),
        ("Objective (Σ squared yield errors)", f"{sv['weighted_sse']:.3e}", f"{ns['weighted_sse']:.3e}"),
        ("Yield RMSE (bp)", f"{sv['yield_rmse_bp']:.2f}", f"{ns['yield_rmse_bp']:.2f}"),
        ("Yield mean absolute error (bp)", f"{sv['yield_mae_bp']:.2f}", f"{ns['yield_mae_bp']:.2f}"),
        ("Largest |yield error| (bp)", f"{sv['yield_max_abs_bp']:.1f}", f"{ns['yield_max_abs_bp']:.1f}"),
        ("Price RMSE (per 100 face)", f"{sv['price_rmse']:.3f}", f"{ns['price_rmse']:.3f}"),
        ("Starting points reaching optimum", f"{sv['n_starts_at_optimum']} / {sv['n_starts']}", f"{ns['n_starts_at_optimum']} / {ns['n_starts']}"),
    ]
    return rows


def selected_rows(results: dict) -> list[tuple[str, str, str, str]]:
    params = np.array([results["svensson"]["params"][k] for k in ("beta0", "beta1", "beta2", "beta3", "tau1", "tau2")])
    t = np.array(SELECTED_MATURITIES, dtype=float)
    z, d, f = svensson_zero(t, *params) * 100, discount_factor(t, params), svensson_forward(t, *params) * 100
    return [(f"{m:g}", f"{zi:.3f}", f"{di:.5f}", f"{fi:.3f}") for m, zi, di, fi in zip(SELECTED_MATURITIES, z, d, f)]


# ------------------------------------------------------------------ renderers
def to_markdown(results: dict) -> str:
    out = [f"# {TITLE}", "", f"_{COURSE} | {AUTHOR} | {_fmt_date(results['quote_date'])} quotes_", "", "## Methodology", ""]
    for s in build_sections(results):
        out += [f"**{s['heading']}.** {_strip(s['html'])}", ""]
        if s.get("display"):
            out += ["$$", s["display"], "$$", ""]
        if s.get("after"):
            out += [_strip(s["after"]), ""]
    out += ["## Results", "", "![Figure 1: zero and forward curves](../output/figures/zero_curve.png)", "",
            "**Table 1. Parameters and fit.**", "", "| | Svensson | Nelson-Siegel |", "|---|---|---|"]
    out += [f"| {_strip(a)} | {b} | {c} |" for a, b, c in parameter_rows(results)]
    out += ["", "**Table 2. Zero curve at selected maturities (continuously compounded).**", "",
            "| t (years) | r(t) % | d(t) | f(t) % |", "|---|---|---|---|"]
    out += [f"| {a} | {b} | {c} | {d} |" for a, b, c, d in selected_rows(results)]
    out += ["", "**Table 3.** The discount rate for every Treasury payment date (245 rows) is in "
            "`output/zero_curve_payment_dates.csv`, on the Curve sheet of the Excel workbook, and in the appendix of the PDF.", ""]
    return "\n".join(out)


def _table(caption: str, head: list[str], rows: list[tuple], cls: str = "") -> str:
    th = "".join(f"<th>{h}</th>" for h in head)
    body = "".join("<tr>" + "".join(f"<td>{c}</td>" for c in r) + "</tr>" for r in rows)
    return f'<table class="{cls}"><caption>{caption}</caption><thead><tr>{th}</tr></thead><tbody>{body}</tbody></table>'


def to_html(results: dict) -> str:
    sections = build_sections(results)
    paras = []
    for s in sections:
        paras.append(f"<p><strong>{s['heading']}.</strong> {s['html']}</p>")
        if s.get("display"):
            paras.append(f'<div class="eq">$$\\begin{{gathered}}{s["display"]}\\end{{gathered}}$$</div>')
        if s.get("after"):
            paras.append(f"<p>{s['after']}</p>")
    pay = results["payment_dates"]
    n_cols, rows_per_col = 3, 46
    per_page = n_cols * rows_per_col
    pages = []
    for start in range(0, len(pay), per_page):
        blocks = []
        for c in range(n_cols):
            chunk = pay[start + c * rows_per_col:start + (c + 1) * rows_per_col]
            if not chunk:
                continue
            rows = [(_fmt_date(r["payment_date"]), f"{r['t_years']:.3f}", f"{r['zero_rate_cc_pct']:.4f}",
                     f"{r['discount_factor']:.5f}", f"{r['inst_forward_cc_pct']:.3f}") for r in chunk]
            blocks.append(_table("", ["Payment date", "t", "r(t) %", "d(t)", "f(t) %"], rows, "dense"))
        pages.append(blocks)
    quote = _fmt_date(results["quote_date"])
    settle = _fmt_date(results["settlement"])
    appendix = ""
    for i, blocks in enumerate(pages):
        head = (f"<h2>Table 3. Discount rate for every Treasury payment date</h2>"
                f"<p class=\"sub\">Union of all coupon and principal dates of the {results['svensson']['n_bonds_total']} bonds: "
                f"{len(pay)} dates. <i>t</i> in years (actual/365) from settlement {settle}; <i>r</i>(<i>t</i>) continuously "
                f"compounded zero rate; <i>d</i>(<i>t</i>) = <i>e</i><sup>−<i>r</i>(<i>t</i>)<i>t</i></sup>; <i>f</i>(<i>t</i>) "
                f"instantaneous forward rate.</p>") if i == 0 else f"<p class=\"sub\">Table 3, continued ({i + 1} of {len(pages)}).</p>"
        appendix += f'<div class="page">{head}<div class="cols">{"".join(blocks)}</div></div>'
    return f"""<!doctype html><html lang="en"><head><meta charset="utf-8"><title>{html_lib.escape(TITLE)}</title>
<link rel="stylesheet" href="{KATEX}/katex.min.css">
<script defer src="{KATEX}/katex.min.js"></script>
<script defer src="{KATEX}/contrib/auto-render.min.js" onload="renderMathInElement(document.body,{{delimiters:[{{left:'$$',right:'$$',display:true}}],throwOnError:false}})"></script>
<style>
@page {{ size: Letter; margin: 0.65in 0.75in; }}
body {{ font-family: -apple-system, "Helvetica Neue", Helvetica, Arial, sans-serif; color: #0b0b0b; font-size: 10.2pt; line-height: 1.36; margin: 0; }}
h1 {{ font-size: 17pt; margin: 0 0 2pt 0; color: #1f3a5f; }}
h2 {{ font-size: 13pt; margin: 0 0 8pt 0; color: #1f3a5f; }}
.sub {{ color: #52514e; font-size: 9.3pt; margin: 0 0 10pt 0; }}
p {{ margin: 0 0 6pt 0; text-align: justify; }}
strong {{ color: #1f3a5f; }}
.eq {{ margin: 2pt 0 6pt 0; font-size: 9.4pt; overflow: hidden; }}
.katex-display {{ margin: 0.25em 0; }}
.page {{ page-break-before: always; }}
figure {{ margin: 0 0 10pt 0; }}
figure img {{ width: 100%; }}
figcaption, caption {{ font-size: 9pt; color: #52514e; text-align: left; margin: 4pt 0 6pt 0; caption-side: top; }}
.two {{ display: flex; gap: 18pt; align-items: flex-start; }}
.two > table {{ flex: 1; }}
table {{ border-collapse: collapse; font-size: 9pt; width: 100%; }}
th, td {{ padding: 2.2pt 6pt; border-bottom: 1px solid #ddd; text-align: right; white-space: nowrap; font-variant-numeric: tabular-nums; }}
th:first-child, td:first-child {{ text-align: left; }}
thead th {{ border-bottom: 1.5px solid #888; color: #333; font-weight: 600; }}
.dense {{ font-size: 7.6pt; }}
.dense th, .dense td {{ padding: 1.1pt 4pt; }}
.cols {{ display: flex; gap: 12pt; align-items: flex-start; }}
.cols > table {{ flex: 1; }}
.note {{ font-size: 8.8pt; color: #52514e; margin-top: 8pt; }}
</style></head><body>
<h1>{html_lib.escape(TITLE)}</h1>
<p class="sub">{html_lib.escape(COURSE)} &nbsp;|&nbsp; {html_lib.escape(AUTHOR)} &nbsp;|&nbsp; WSJ Treasury quotes of {quote}, settlement {settle}</p>
{''.join(paras)}

<div class="page">
<h2>Results</h2>
<figure><img src="{FIGURE}" alt="Fitted zero-coupon and forward curves with Treasury street yields">
<figcaption><b>Figure 1.</b> Fitted continuously compounded zero curve <i>r</i>(<i>t</i>) and instantaneous forward curve <i>f</i>(<i>t</i>) against the street yields of the {results['svensson']['n_bonds_in_fit']} bonds in the fit (dots) and the {results['svensson']['n_bonds_total'] - results['svensson']['n_bonds_in_fit']} excluded bonds under 3 months (hollow circles). {quote} quotes, settlement {settle}.</figcaption></figure>
<div class="two">
{_table("<b>Table 1.</b> Parameters and fit quality (bonds in fit).", ["", "Svensson", "Nelson-Siegel"], parameter_rows(results))}
{_table("<b>Table 2.</b> Zero curve at selected maturities, continuously compounded.", ["t (years)", "r(t) %", "d(t)", "f(t) %"], selected_rows(results))}
</div>
<p class="note">All bonds and the full payment-date table are also in the Excel workbook (sheets Bonds and Curve) and in the repository's CSV files.</p>
</div>

{appendix}
</body></html>"""


CHROME_CANDIDATES = (
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    "google-chrome", "chromium", "chrome",
)


def print_pdf(html_path: Path, pdf_path: Path) -> None:
    """Print the HTML document to PDF with headless Chrome, if available."""
    chrome = next((c for c in CHROME_CANDIDATES if Path(c).exists() or shutil.which(c)), None)
    if chrome is None:
        print("Chrome not found: PDF not regenerated (open the HTML and print to PDF instead)")
        return
    subprocess.run(
        [chrome, "--headless=new", "--disable-gpu", "--no-pdf-header-footer", "--virtual-time-budget=15000",
         f"--print-to-pdf={pdf_path.resolve()}", html_path.resolve().as_uri()],
        check=True, capture_output=True,
    )
    print(f"PDF written to {pdf_path.resolve()}")


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--results", default="../output/results.json")
    parser.add_argument("--out-dir", default="../writeup")
    parser.add_argument("--no-pdf", action="store_true", help="skip the headless-Chrome PDF step")
    args = parser.parse_args(argv)
    results = json.loads(Path(args.results).read_text())
    out = Path(args.out_dir)
    out.mkdir(parents=True, exist_ok=True)
    (out / "Miniproject2_ZCB_Term_Structure_Writeup.md").write_text(to_markdown(results))
    html_path = out / "Miniproject2_ZCB_Term_Structure_Writeup.html"
    html_path.write_text(to_html(results))
    words = sum(len(t.split()) for _, t in build_paragraphs(results))
    print(f"Write-up written to {out.resolve()} (methodology {words} words)")
    if not args.no_pdf:
        print_pdf(html_path, out / "Miniproject2_ZCB_Term_Structure_Writeup.pdf")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
