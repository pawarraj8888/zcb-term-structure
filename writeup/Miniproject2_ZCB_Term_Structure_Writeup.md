# Miniproject 2 - Creating a ZCB Term Structure

_FRE 6103 Valuation for Financial Engineering, NYU Tandon | Raj Pawar | Sep 4, 2026 quotes_

## Methodology

**Data and conventions.** Wall Street Journal (Tullett Prebon) 3 pm quotes for all 353 US Treasury notes and bonds on Friday, September 4, 2026, the Week 2 data file. Quotes are clean prices in 32nds with a third decimal in eighths of a 32nd (99.256 = 99 + 25.75/32); we price off the asked side. Regular-way settlement is T+1 business days, Tuesday, September 8, 2026 (Monday is Labor Day). Both conventions were checked by recomputing every WSJ asked yield: median discrepancy 0.02 bp and RMSE 0.5 bp, against 1.8 bp for the next-best settlement date and over 86 bp if the quotes are read as decimals.

**Cash flows.** Each bond pays coupon/2 semi-annually on its maturity day-of-month (month-end maturities pay on month-ends) and 100 at maturity. Accrued interest is actual/actual within the coupon period, coupon/2 × days accrued / days in period; dirty price = clean price + accrued (slide 34). Time to each payment is t = days/365 from settlement. Street yields follow the Treasury convention (semi-annual compounding with a fractional first period, simple interest when one payment remains), which reproduces Excel's YIELD.

**Model.** The continuously compounded zero rate follows the Svensson (1994) form suggested with the data file, the discount factor is d(t) = e^−r(t)t, and each bond's model dirty price is the sumproduct of its payment stream with the discount factors (slide 33):

$$
r(t)=\beta_0+\beta_1\,\frac{1-e^{-t/\tau_1}}{t/\tau_1}+\beta_2\left(\frac{1-e^{-t/\tau_1}}{t/\tau_1}-e^{-t/\tau_1}\right)+\beta_3\left(\frac{1-e^{-t/\tau_2}}{t/\tau_2}-e^{-t/\tau_2}\right)\\[2pt] d(t)=e^{-r(t)\,t},\qquad P_i^{\mathrm{model}}=\sum_k CF_{ik}\,d(t_{ik})
$$

The form nests Nelson-Siegel (β_3 = 0), is smooth, has finite limits at both ends (r(0) = β_0 + β_1, r(∞) = β_0) and gives the instantaneous forward curve in closed form, f(t) = β_0 + β_1e^−t/τ_1 + β_2(t/τ_1)e^−t/τ_1 + β_3(t/τ_2)e^−t/τ_2.

**Estimation.** The six parameters minimise the sum of squared price errors, each divided by dirty price times modified duration D_i, which is the first-order yield error; the objective is therefore the sum of squared yield errors (slide 31) while every model price stays a linear sumproduct:

$$
\min_{\beta,\tau}\;\sum_{i\in\text{fit}}\left(\frac{P_i^{\mathrm{model}}-P_i^{\mathrm{market}}}{P_i^{\mathrm{market}}\,D_i}\right)^2,\qquad \beta_0\ge 0,\;\tau_1,\tau_2\ge 0.05
$$

Bonds with under 3 months to maturity (14 of 353) are excluded from estimation, following Gürkaynak, Sack and Wright (2007): their prices reflect money-market conditions and, with near-zero duration, a one-cent quote error is a 50 bp yield error; they are still priced off the curve as an out-of-sample check. The objective is multimodal in (τ_1, τ_2), so the least-squares solver is started from 19 points on a τ grid; 11 reach the same optimum. In Excel the same objective is the sum-of-squares cell on the Curve sheet, minimised with Solver (GRG Nonlinear); the workbook reproduces every Python model price, yield and zero rate.

**Results.** β_0 = 0.0000 (bound active), β_1 = 0.0363, β_2 = 0.0447, β_3 = 0.1622, τ_1 = 1.76, τ_2 = 15.93 (Table 1). The curve prices the 339 bonds with a yield RMSE of 2.9 bp (mean absolute 2.1 bp; worst 14.8 bp, the 4.375% Dec 15, 2026 note); Nelson-Siegel gives 7.5 bp, so the second hump is needed. The zero curve rises from 3.63% at the short end to 4.14% at 1 year, 4.49% at 5, 4.79% at 10, 5.33% at 20 and 5.31% at 30 years, peaking at 5.38% near 24 years (Figure 1, Table 2). Table 3 lists r(t), d(t) and f(t) for all 245 distinct Treasury payment dates from Sep 15, 2026 to Aug 15, 2056. β_0, the rate at infinite maturity, is weakly identified by a 30-year sample: pinning it anywhere from 0% to 3% worsens the fit by only 0.3 bp RMSE and moves the in-sample curve by about 10 bp at the two ends (3% in d(30y)), so the payment-date discount factors are robust to it; with β_0 = 0 the forward rate decays beyond the sample, so the curve should not be extrapolated past 30 years.

## Results

![Figure 1: zero and forward curves](../output/figures/zero_curve.png)

**Table 1. Parameters and fit.**

| | Svensson | Nelson-Siegel |
|---|---|---|
| β_0 | 0.00000 | 0.05879 |
| β_1 | 0.03628 | -0.01870 |
| β_2 | 0.04467 | -0.00000 |
| β_3 | 0.16218 | 0.00000 |
| τ_1 (years) | 1.761 | 7.015 |
| τ_2 (years) | 15.927 | — |
| Bonds in fit | 339 | 339 |
| Objective (Σ squared yield errors) | 2.810e-05 | 1.922e-04 |
| Yield RMSE (bp) | 2.88 | 7.52 |
| Yield mean absolute error (bp) | 2.06 | 5.85 |
| Largest |yield error| (bp) | 14.8 | 38.9 |
| Price RMSE (per 100 face) | 0.185 | 0.382 |
| Starting points reaching optimum | 11 / 19 | 4 / 4 |

**Table 2. Zero curve at selected maturities (continuously compounded).**

| t (years) | r(t) % | d(t) | f(t) % |
|---|---|---|---|
| 0.25 | 3.796 | 0.99055 | 3.948 |
| 0.5 | 3.933 | 0.98053 | 4.179 |
| 1 | 4.133 | 0.95951 | 4.450 |
| 2 | 4.340 | 0.91685 | 4.591 |
| 3 | 4.422 | 0.87575 | 4.577 |
| 5 | 4.494 | 0.79875 | 4.674 |
| 7 | 4.588 | 0.72530 | 4.995 |
| 10 | 4.794 | 0.61918 | 5.534 |
| 15 | 5.131 | 0.46315 | 5.964 |
| 20 | 5.329 | 0.34445 | 5.802 |
| 25 | 5.377 | 0.26073 | 5.298 |
| 30 | 5.310 | 0.20329 | 4.645 |

**Table 3.** The discount rate for every Treasury payment date (245 rows) is in `output/zero_curve_payment_dates.csv`, on the Curve sheet of the Excel workbook, and in the appendix of the PDF.
