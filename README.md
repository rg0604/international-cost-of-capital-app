# International Cost of Capital Explorer

A Python 3.12 Streamlit app for comparing a project's USD and foreign-currency cost of equity, borrowing cost, and WACC. It follows Professor Sergio Rebelo's *International Valuation* (2026) lecture slides 573–587, including the Colombia example.

## Quick start

```bash
python3.12 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
streamlit run app.py
```

On Windows PowerShell, activate using `.venv\Scripts\Activate.ps1`.

To deploy on Streamlit Community Cloud: commit **app.py**, **requirements.txt**, and **README.md** at the root of a GitHub repository; choose the `main` branch and `app.py`, set Python **3.12** in advanced settings, and leave Secrets empty.

## Defaults

- Industry: **Software (System & Application)**
- Target: **United States**
- Tax: **U.S. corporate tax rate**, applied to both financing scenarios
- Equity risk premium: **6.5%**
- Target debt share: **30%**

Select an industry and a target country, or manually override beta, tax rates, the USD BBB spread, local BBB spread, and both government yields. The app displays costs side by side and supports two CSV exports (selected project and all-industries if the Damodaran spreadsheet loads).

## Exact lecture methodology

All rates in formulas are decimal rates (for instance 6.5% is 0.065). Let `d` be target debt share and `τ` the **single selected project tax rate**. The app assumes market-value weights and full deductibility of interest.

1. `D/E = d/(1-d)`
2. `βE = βU × [1 + (1-τ) × D/E]`
3. `Re,USD = Rf,USD + βE × ERP`
4. `Rd,USD = Rf,USD + sUSD`
5. `F = (1+Rf,foreign)/(1+Rf,USD)`
6. `Re,foreign = (1+Re,USD)×F - 1`
7. `sforeign = sUSD×F` unless a manual local spread is supplied
8. `Rd,foreign = Rf,foreign + sforeign`
9. `WACC,currency = (1-d)×Re,currency + d×Rd,currency×(1-τ)`

This is gross-return currency conversion using government yields, **not** an additive yield adjustment. Do not add a separate country equity risk premium. Government yields are only risk-free proxies. Risk is measured relative to the U.S. market, and currency conversion is motivated by relative PPP and comparable real rates.

## Sources and provenance

- Damodaran U.S. unlevered industry betas: https://pages.stern.nyu.edu/~adamodar/pc/datasets/betas.xls. Downloaded on demand. If blocked or unparsable, beta must be manually entered; illustrative 0.80 comes from the **Colombia classroom example** and is **not** a Software industry quote.
- Damodaran large-nonfinancial Baa2/BBB default spread: https://pages.stern.nyu.edu/~adamodar/New_Home_Page/datafile/ratings.html. The source page currently indicates January 2026 and a 1.11% Baa2/BBB spread, but the app fetches it rather than embedding an undated live quote. If blocked, illustrative **1.50%** comes from the Colombia classroom example (not a recent market quote).
- OECD corporate tax statistics: https://www.oecd.org/en/publications/corporate-tax-statistics-2026_73af6222-en/full-report/statutory-corporate-income-tax-rates_ce84abb9.html. Best-effort HTML table retrieval; dated **2026** snapshot from the professor's lecture (slide 550) serves as fallback where the website does not expose a table.
- Trading Economics public 10-year government yields: https://tradingeconomics.com/bonds. Best-effort retrieval from server-rendered HTML. Dated **November 6, 2026** lecture snapshot (slide 575) serves as fallback for 16 countries, and all remaining countries require user-supplied yields. Trading Economics may block web scraping; this app makes no guarantee of real-time quotes or unrestricted access.

**Data status and attempted fetch timestamp are displayed inside the app.** Refresh data retries the network sources and invalidates cached responses. Manual overrides are visibly identified. No API key or secret is required.

## Verification

### Lecture's Colombia example (slides 578–585)

Input: `βU=0.8`, `d=0.30`, `τ=0.35`, `ERP=0.065`, `Rf,USD=0.04`, `Rf,COP=0.12`, `sUSD=0.015`.

- `D/E = 0.30/0.70 = 0.4285714`
- `βE = 0.8×(1+0.65×0.30/0.70) = 1.0228571`
- `Re,USD = 0.04+1.0228571×0.065 = 10.6486%`
- `Rd,USD = 4.0%+1.5% = 5.5%`
- `WACC,USD = 0.7×10.648571%+0.3×5.5%×0.65 = 8.5265%`
- `F = 1.12/1.04 = 1.0769231`
- `Re,COP = (1+Re,USD)×F-1 = 19.1600%`
- `sCOP = 1.5%×F = 1.6154%`
- `Rd,COP = 12%+1.6154% = 13.6154%`
- `WACC,COP = 0.7×19.1600%+0.3×13.6154%×0.65 = 16.0670%`

### Required invariants

1. With identical USD and foreign government yields and no foreign spread override, `F=1`, so all costs in both currencies are equal.
2. With `d=0`, WACC equals cost of equity in each currency.
3. Changing only the foreign BBB spread override changes only foreign pretax debt cost / foreign WACC, not USD outputs.

## Limitations

Source websites are dynamic and may restrict automated requests. You should confirm source dates, check each input's provenance, and independently verify any real-world market quote before relying on it. OECD statutory rates may differ from the effective project tax rate. Countries without an automatically available yield require a manual override. The app does **not** fabricate missing market rates.

## Development record

Model: **OpenAI GPT-6**. Initial build: **November 15, 2026**.

Development prompt summary: *Create a complete Python 3.12 Streamlit international cost of capital comparison application using the exact formulas from the uploaded Sergio Rebelo 2026 lecture; obtain current Damodaran unlevered industry betas and large-nonfinancial BBB default spread, OECD statutory corporate taxes, and Trading Economics 10-year government yields where publicly accessible; include honest failure handling, dated lecture fallbacks, manual overrides, side-by-side cost of equity/debt/WACC, all-industry comparison, CSV exports, and explicit verification of the Colombia example and three requested invariants. Supply app.py, requirements.txt, and README.md.*
