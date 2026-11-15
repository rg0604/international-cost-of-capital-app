"""International Cost of Capital Explorer — Streamlit / Python 3.12.

Methodology: Sergio Rebelo, International Valuation, lecture slides 573–587 (2026).
All calculation inputs are proportions internally (not percentage points).
"""
from __future__ import annotations

from datetime import datetime, timezone
from io import BytesIO
import re
from typing import Any

import altair as alt
import pandas as pd
import pycountry
import requests
import streamlit as st
from bs4 import BeautifulSoup

st.set_page_config(page_title="International Cost of Capital", page_icon="🌐", layout="wide")

BETA_URL = "https://pages.stern.nyu.edu/~adamodar/pc/datasets/betas.xls"
SPREAD_URL = "https://pages.stern.nyu.edu/~adamodar/New_Home_Page/datafile/ratings.html"
TAX_URL = "https://www.oecd.org/en/publications/corporate-tax-statistics-2026_73af6222-en/full-report/statutory-corporate-income-tax-rates_ce84abb9.html"
TE_URL = "https://tradingeconomics.com/bonds"
LECTURE_DATE = "2026-11-06"
US_NAME = "United States"
DEFAULT_INDUSTRY = "Software (System & Application)"
HEADERS = {"User-Agent": "Mozilla/5.0 (compatible; educational-streamlit-app/1.0)", "Accept": "text/html,application/vnd.ms-excel,*/*"}

# Explicitly dated FALLBACK observations from professor's 2026 slides, not live quotes.
# Slide 575: 10-year yields on Nov 6, 2026. No value is invented for other countries.
LECTURE_YIELDS = {
    "Brazil": 15.2, "Colombia": 12.8, "Mexico": 9.1, "India": 7.2,
    "Poland": 6.2, "United Kingdom": 5.2, "Australia": 5.0,
    "United States": 4.9, "France": 4.8, "Slovakia": 4.2,
    "Portugal": 3.7, "Canada": 3.7, "Germany": 3.3,
    "Japan": 2.6, "China": 1.6, "Switzerland": 0.4,
}
# OECD 2026 combined statutory rates as reproduced in professor's slide 550.
# These are a snapshot, not a claim of a live OECD web fetch.
LECTURE_TAX = {
    "France":36.1,"Argentina":35.0,"Germany":30.1,"Australia":30.0,
    "Kenya":30.0,"Mexico":30.0,"Nigeria":30.0,"Senegal":30.0,"Zambia":30.0,
    "Japan":29.7,"Portugal":29.5,"Peru":29.5,"Pakistan":29.0,
    "Republic of the Congo":28.0,"New Zealand":28.0,"Italy":27.8,"South Korea":27.5,
    "Chile":27.0,"Dominican Republic":27.0,"South Africa":27.0,
    "Greenland":26.5,"Canada":26.0,"Netherlands":25.8,"United States":25.5,
    "India":25.2,"Angola":25.0,"Belgium":25.0,"China":25.0,"Spain":25.0,
    "United Kingdom":25.0,"Monaco":25.0,"Panama":25.0,"Philippines":25.0,
    "Uruguay":25.0,"Turkey":25.0,"Malaysia":24.0,"Slovakia":24.0,
    "Luxembourg":23.9,"Austria":23.0,"Israel":23.0,"Egypt":22.5,
    "Botswana":22.0,"Denmark":22.0,"Greece":22.0,"Indonesia":22.0,
    "Norway":22.0,"Slovenia":22.0,"Estonia":22.0,"Czechia":21.0,
    "Sweden":20.6,"Cabo Verde":20.4,"Iceland":20.0,"Finland":20.0,
    "Jordan":20.0,"Kazakhstan":20.0,"Latvia":20.0,"Saudi Arabia":20.0,
    "Thailand":20.0,"Vietnam":20.0,"Switzerland":19.5,"Poland":19.0,
    "Brunei":18.5,"Armenia":18.0,"Croatia":18.0,"Ukraine":18.0,
    "Singapore":17.0,"Lithuania":17.0,"Hong Kong":16.5,"Romania":16.0,
    "Albania":15.0,"Georgia":15.0,"Kuwait":15.0,"Ireland":12.5,
    "Macau":12.0,"Bulgaria":10.0,"Paraguay":10.0,"Qatar":10.0,
    "United Arab Emirates":9.0,"Hungary":9.0,
}

ALIASES = {
    "united states":"United States","united states of america":"United States","usa":"United States","us":"United States",
    "united kingdom":"United Kingdom","great britain":"United Kingdom","uk":"United Kingdom",
    "turkiye":"Turkey","türkiye":"Turkey","korea, republic of":"South Korea","south korea":"South Korea",
    "republic of korea":"South Korea","czech republic":"Czechia","czechia":"Czechia",
    "slovak republic":"Slovakia","uae":"United Arab Emirates", "hong kong sar, china":"Hong Kong",
    "hong kong":"Hong Kong", "brunei darussalam":"Brunei", "viet nam":"Vietnam",
    "russian federation":"Russia", "china":"China", "colombia":"Colombia",
}

def normalized_country(name: str) -> str:
    key = re.sub(r"\s+", " ", str(name)).strip().lower()
    return ALIASES.get(key, str(name).strip())


def country_options() -> list[str]:
    names = {normalized_country(c.name) for c in pycountry.countries}
    names.update(LECTURE_TAX)
    names.update(LECTURE_YIELDS)
    names.add(US_NAME)
    return sorted(names)


def download(url: str) -> requests.Response:
    r = requests.get(url, headers=HEADERS, timeout=(8, 22))
    r.raise_for_status()
    return r


def strip_column(x: Any) -> str:
    if isinstance(x, tuple):
        x = " ".join(str(p) for p in x if not str(p).startswith("Unnamed"))
    return re.sub(r"\s+", " ", str(x)).strip()


def pct_number(value: Any) -> float | None:
    """Parse a percent as percent POINTS, e.g. '1.11%' -> 1.11."""
    if value is None or pd.isna(value):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip().replace(",", "").replace("−", "-")
    hit = re.search(r"[-+]?\d+(?:\.\d+)?", s)
    return float(hit.group()) if hit else None


def beta_number(value: Any) -> float | None:
    try:
        result = float(str(value).strip().replace(",", ""))
        return result if 0 <= result < 20 else None
    except (TypeError, ValueError):
        return None


def parse_beta_workbook(payload: bytes) -> pd.DataFrame:
    """Find Damodaran's actual industry/unlevered-beta header despite workbook layout changes."""
    workbook = pd.ExcelFile(BytesIO(payload), engine="xlrd")
    for sheet in workbook.sheet_names:
        raw = pd.read_excel(workbook, sheet_name=sheet, header=None)
        for i in range(min(len(raw), 22)):
            line = [strip_column(x).lower() for x in raw.iloc[i].tolist()]
            inds = [j for j, v in enumerate(line) if "industry" in v and "name" in v]
            beta_cols = [j for j, v in enumerate(line) if "unlevered beta" in v and "cash" not in v]
            if not beta_cols:
                beta_cols = [j for j, v in enumerate(line) if "unlevered" in v and "beta" in v and "cash" not in v]
            if not inds or not beta_cols:
                continue
            names_col, beta_col = inds[0], beta_cols[0]
            rows = raw.iloc[i+1:, [names_col, beta_col]].copy()
            rows.columns = ["industry", "beta_u"]
            rows["industry"] = rows["industry"].astype(str).str.strip()
            rows["beta_u"] = rows["beta_u"].map(beta_number)
            rows = rows.dropna(subset=["beta_u"])
            rows = rows[~rows["industry"].str.contains("total|overall|aggregate|all firms", case=False, na=False)]
            rows = rows[rows["industry"].str.lower().ne("nan")]
            if len(rows) >= 12:
                return rows.drop_duplicates("industry").sort_values("industry").reset_index(drop=True)
    raise ValueError("Could not identify industry and unlevered beta columns in Damodaran workbook")


def fetch_spread() -> tuple[float, str]:
    """Extract the Baa2/BBB large non-financial default spread, in pp."""
    html = download(SPREAD_URL).text
    for table in pd.read_html(BytesIO(html.encode()), flavor="lxml"):
        rows = table.astype(str)
        for _, row in rows.iterrows():
            vals = row.tolist()
            if any(re.search(r"Baa2\s*/\s*BBB\b", v, re.I) for v in vals):
                idx = next(i for i,v in enumerate(vals) if re.search(r"Baa2\s*/\s*BBB\b",v,re.I))
                # Rating/spread columns appear adjacent. First spread after the first rating.
                for v in vals[idx+1:idx+3]:
                    if "%" in v:
                        n = pct_number(v)
                        if n is not None:
                            date_match = re.search(r"(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+20\d{2}",html,re.I)
                            return n, (date_match.group() if date_match else "date not stated")
    # HTML table layout fallback: bounded regex after Baa2/BBB.
    m = re.search(r"Baa2\s*/\s*BBB[\s\S]{0,400}?(\d+(?:\.\d+)?)\s*%", html, re.I)
    if m:
        return float(m.group(1)), "source-page date; verify"
    raise ValueError("Baa2/BBB default spread not found on Damodaran's page")


def scrape_oecd_tax_table() -> dict[str, float]:
    """Best effort: extract combined statutory tax rates from OECD published HTML tables."""
    html = download(TAX_URL).text
    tables = pd.read_html(BytesIO(html.encode()), flavor="lxml")
    result: dict[str, float] = {}
    for tb in tables:
        cols = [strip_column(c).lower() for c in tb.columns]
        country_cols = [i for i, c in enumerate(cols) if "country" in c or "jurisdiction" in c]
        rate_cols = [i for i, c in enumerate(cols) if ("combined" in c or "statutory" in c) and ("rate" in c or "2026" in c)]
        if not country_cols or not rate_cols:
            continue
        for _, row in tb.iterrows():
            name = normalized_country(str(row.iloc[country_cols[0]]))
            val = pct_number(row.iloc[rate_cols[0]])
            if val is not None and 0 <= val <= 65:
                result[name] = val
    if not result:
        raise ValueError("OECD web page provides no machine-readable combined tax-rate table")
    return result


def parse_te_bond_page() -> dict[str,float]:
    """Best effort public HTML scrape. Trading Economics can change/disable its markup."""
    html = download(TE_URL).text
    result: dict[str,float] = {}
    for table in pd.read_html(BytesIO(html.encode()), flavor="lxml"):
        cols = [strip_column(c).lower() for c in table.columns]
        name_idx = next((i for i,c in enumerate(cols) if "country" in c), None)
        y_idx = next((i for i,c in enumerate(cols) if c in ("yield", "last", "10y", "10-year yield") or "10y yield" in c), None)
        if name_idx is None or y_idx is None:
            continue
        for _,row in table.iterrows():
            name = normalized_country(str(row.iloc[name_idx]))
            val = pct_number(row.iloc[y_idx])
            if val is not None and -5 < val < 100:
                result[name] = val
    if not result:
        raise ValueError("Trading Economics bond table not present in server-rendered HTML")
    return result


@st.cache_data(ttl=60*60*12, show_spinner=False)
def source_bundle(refresh_token: int) -> dict[str,Any]:
    """Independent sources: failures never turn example values into purported live data."""
    out: dict[str,Any] = {"beta":pd.DataFrame(columns=["industry","beta_u"]), "beta_status":"Unavailable",
        "spread":None,"spread_status":"Unavailable","tax":LECTURE_TAX.copy(),
        "tax_status":"OECD Corporate Tax Statistics 2026; lecture slide 550 (snapshot)",
        "yields":LECTURE_YIELDS.copy(),
        "yield_status":"Trading Economics, 2026-11-06; lecture slide 575 (snapshot)", "errors":[],
        "fetched_at":datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")}
    try:
        payload = download(BETA_URL).content
        out["beta"] = parse_beta_workbook(payload)
        out["beta_status"] = "Damodaran, latest downloaded workbook; date in source workbook"
    except Exception as e:
        out["errors"].append(f"Industry betas: {str(e)[:175]}")
    try:
        out["spread"],date = fetch_spread()
        out["spread_status"] = f"Damodaran Baa2/BBB (large non-financial), {date}"
    except Exception as e:
        out["errors"].append(f"BBB borrowing spread: {str(e)[:175]}")
    try:
        online_tax = scrape_oecd_tax_table()
        # Only update reliably recognized countries; other snapshot countries stay marked as snapshot.
        if len(online_tax) >= 15:
            out["tax"].update(online_tax)
            out["tax_status"] = "OECD web table (where parsed); remaining countries: 2026 lecture snapshot"
    except Exception as e:
        out["errors"].append(f"OECD online tax table: {str(e)[:175]}")
    try:
        online_yields = parse_te_bond_page()
        if len(online_yields) >= 3:
            out["yields"].update(online_yields)
            out["yield_status"] = "Trading Economics live public bond table (where parsed); remaining countries: Nov 6, 2026 lecture snapshot"
    except Exception as e:
        out["errors"].append(f"Trading Economics public yields: {str(e)[:175]}")
    return out


def compute(beta_u: float, d: float, tax: float, erp: float,
            rf_us: float, rf_local: float, spread_us: float,
            spread_local_override: float | None = None) -> dict[str,float]:
    """Slide 579–585 equations. Inputs are rates in decimals (0.065 = 6.5%)."""
    if not (0 <= d < 1):
        raise ValueError("Target debt share must be between 0 and less than 100%")
    if not (0 <= tax <= 1):
        raise ValueError("Corporate tax must be between 0% and 100%")
    if rf_us <= -1 or rf_local <= -1:
        raise ValueError("Government yield must exceed -100%")
    de = d/(1-d)
    beta_e = beta_u*(1+(1-tax)*de)
    re_us = rf_us+beta_e*erp
    rd_us = rf_us+spread_us
    wacc_us = (1-d)*re_us+d*rd_us*(1-tax)
    factor = (1+rf_local)/(1+rf_us)
    re_local = (1+re_us)*factor-1
    local_spread = spread_us*factor if spread_local_override is None else spread_local_override
    rd_local = rf_local+local_spread
    wacc_local = (1-d)*re_local+d*rd_local*(1-tax)
    return {"beta_u":beta_u,"beta_e":beta_e,"d_e":de,"tax":tax,"d":d,
            "rf_us":rf_us,"rf_local":rf_local,"erp":erp,"spread_us":spread_us,
            "spread_local":local_spread,"fx_factor":factor,
            "re_us":re_us,"rd_us":rd_us,"wacc_us":wacc_us,
            "re_local":re_local,"rd_local":rd_local,"wacc_local":wacc_local}


def rate(n:float) -> str:
    return f"{n*100:.4f}%"


def data_source_label(name: str, online: dict, base: dict, status: str) -> str:
    return status if name in online else base.get(name, "Not available")


def main() -> None:
    """Lecture-slide-first layout, with detailed manual overrides folded away."""
    st.markdown("""<style>
      .block-container {max-width: 1160px; padding-top: 2.0rem;}
      [data-testid="stSidebar"] {background: #f7f7f7;}
      [data-testid="stSidebar"] .block-container {padding-top: 1.4rem;}
      h1 {font-size: 2.1rem !important; letter-spacing: -.035em;}
      h2 {font-size: 1.3rem !important; letter-spacing: -.02em;}
      h3 {font-size: 1.08rem !important;}
      [data-testid="stMetricValue"] {font-size: 1.75rem;}
      [data-testid="stMetricLabel"] {font-size: .83rem;}
      div[data-testid="stVerticalBlock"] > div:has(> [data-testid="stMetric"]) {padding-top: .3rem;}
      hr {margin: .7rem 0 1.2rem;}
    </style>""", unsafe_allow_html=True)
    if "refresh_counter" not in st.session_state:
        st.session_state.refresh_counter = 0
    with st.spinner("Loading source data…"):
        info = source_bundle(st.session_state.refresh_counter)
    industry_frame = info["beta"]
    industries = sorted(set(industry_frame["industry"].tolist() + [DEFAULT_INDUSTRY]))
    country_list = country_options()

    with st.sidebar:
        st.subheader("Project assumptions")
        if st.button("↻ Refresh sources data", use_container_width=True):
            st.session_state.refresh_counter += 1
            st.rerun()
        industry = st.selectbox("U.S. industry", industries, index=industries.index(DEFAULT_INDUSTRY))
        st.caption("Damodaran's U.S. industry unlevered beta")
        exact = industry_frame[industry_frame["industry"].str.casefold() == industry.casefold()]
        if exact.empty and industry == DEFAULT_INDUSTRY and not industry_frame.empty:
            exact = industry_frame[industry_frame["industry"].str.contains(r"software.*(system|application)", case=False, regex=True)]
        source_beta = float(exact.iloc[0]["beta_u"]) if not exact.empty else None
        if source_beta is None:
            st.warning("Industry beta unavailable; enter a manual value below.")
        beta_override = st.checkbox("Enter unlevered beta manually", value=source_beta is None)
        if beta_override:
            beta_u = st.number_input("Unlevered beta", 0.0, 10.0,
                 value=float(source_beta if source_beta is not None else 0.80), step=0.05)
        else:
            beta_u = source_beta
        st.number_input("Annual equity risk premium (%)", min_value=0.0, max_value=50.0,
                        value=6.5, step=0.25, key="erp_pp")
        st.slider("Fraction financed with debt (%)", 0, 95, value=30, key="debt_pp")
        spread = info["spread"]
        spread_override = st.checkbox("Override BBB borrowing spread", value=spread is None)
        if spread_override:
            usd_spread_pp = st.number_input("USD borrowing spread (%)", 0.0, 40.0,
                    value=float(spread if spread is not None else 1.5), step=0.05)
        else:
            usd_spread_pp = spread
        if spread is None:
            st.caption("1.50% is only the lecture illustration, not a current market quote.")
        country = st.selectbox("Project country / currency", country_list,
                      index=country_list.index(US_NAME))
        use_local_tax = st.checkbox("Override corporate tax rate selection: use target country", value=False)
        st.caption("Unselected: U.S. corporate tax rate. One project tax rate applies in both currencies.")

    # Inputs shared between the two currency calculations
    tax_us = info["tax"].get(US_NAME)
    tax_local = info["tax"].get(country)
    rf_us = info["yields"].get(US_NAME)
    rf_local = info["yields"].get(country)

    st.title("Cost of capital")
    st.caption("U.S. industry betas, corporate taxes, BBB borrowing spreads, and 10-year government bond yields")
    st.subheader("10-year yields used in the calculation")
    ycol1, ycol2 = st.columns(2, gap="large")
    with ycol1:
        us_yield_manual = st.checkbox("Enter U.S. 10-year yield manually", value=rf_us is None)
        if us_yield_manual:
            us_yield_pp = st.number_input("U.S. 10-year yield (%)", min_value=-10.0, max_value=60.0,
                    value=float(rf_us if rf_us is not None else 4.9), step=0.10)
        else:
            us_yield_pp = rf_us
        st.caption(f"Source: {info['yield_status']}")
    with ycol2:
        st.markdown(f"**Selected country: {country}**")
        if country == US_NAME:
            st.caption("The selected country is the United States (USD). No currency conversion is needed.")
            local_yield_pp = us_yield_pp
            local_yield_manual = False
        else:
            local_yield_manual = st.checkbox("Enter selected-country yield manually", value=rf_local is None)
            if local_yield_manual:
                local_yield_pp = st.number_input(f"{country} 10-year yield (%)", min_value=-10.0,
                      max_value=100.0, value=float(rf_local if rf_local is not None else 0.0), step=0.10)
            else:
                local_yield_pp = rf_local
            st.caption(f"Source: {info['yield_status']}")
    if us_yield_pp is None or local_yield_pp is None:
        st.error("Yield is missing. Enable the appropriate manual yield input.")
        st.stop()

    # Tax overrides are required by the assignment but belong in secondary settings.
    with st.expander("Advanced inputs: tax rates and foreign-currency borrowing spread"):
        a,b = st.columns(2)
        with a:
            tax_us_override = st.checkbox("Override U.S. corporate tax rate", value=tax_us is None)
            if tax_us_override:
                tax_us_pp = st.number_input("U.S. tax rate (%)", 0.0, 65.0,
                              value=float(tax_us if tax_us is not None else 25.5), step=0.1)
            else:
                tax_us_pp = tax_us
        with b:
            tax_local_override = st.checkbox("Override selected-country corporate tax rate", value=False)
            if tax_local_override:
                tax_local_pp = st.number_input(f"{country} tax rate (%)", 0.0, 65.0,
                              value=float(tax_local if tax_local is not None else 25.0), step=0.1)
            else:
                tax_local_pp = tax_local
        fx_factor = (1 + local_yield_pp/100)/(1 + us_yield_pp/100)
        auto_local_spread = usd_spread_pp * fx_factor
        foreign_override = st.checkbox("Override foreign-currency BBB borrowing spread", value=False,
                                     disabled=(country == US_NAME))
        if foreign_override:
            local_spread_pp = st.number_input("Foreign-currency borrowing spread (%)", 0.0, 40.0,
                                        value=float(auto_local_spread), step=0.05)
        else:
            local_spread_pp = auto_local_spread
        st.caption("Default foreign spread is USD BBB spread × (1 + foreign 10-year yield) / (1 + U.S. 10-year yield).")
    tax_pp = tax_local_pp if use_local_tax else tax_us_pp
    if tax_pp is None:
        st.error("The chosen jurisdiction's corporate tax rate is unavailable. Enable its manual override in Advanced inputs.")
        st.stop()
    if usd_spread_pp is None:
        st.error("The USD BBB spread could not be retrieved; enable its manual override in the sidebar.")
        st.stop()

    result = compute(beta_u, st.session_state.debt_pp / 100, tax_pp/100,
            st.session_state.erp_pp/100, us_yield_pp/100, local_yield_pp/100,
            usd_spread_pp/100, local_spread_pp/100 if foreign_override else None)

    st.subheader("Borrowing spreads")
    sc1, sc2 = st.columns(2)
    sc1.metric("Borrowing spread · USD", f"{usd_spread_pp:.4f}%")
    sc2.metric(f"Borrowing spread · {country}", f"{result['spread_local']*100:.4f}%")
    st.caption(f"USD credit benchmark: {info['spread_status']}. The foreign spread is {'manually overridden' if foreign_override else 'converted using the government-yield factor'}. ")
    st.divider()

    st.subheader(industry)
    st.caption(f"{country} · U.S. industry beta data: {info['beta_status']}. Applied corporate tax: {tax_pp:.2f}% ({'selected country' if use_local_tax else 'United States'}).")
    m1,m2,m3,m4 = st.columns(4)
    m1.metric("U.S. 10-year yield", f"{us_yield_pp:.3f}%")
    m2.metric("Selected 10-year yield", f"{local_yield_pp:.3f}%")
    m3.metric("Corporate tax rate", f"{tax_pp:.2f}%")
    m4.metric("Unlevered beta", f"{beta_u:.4f}")
    st.divider()
    left,right = st.columns(2, gap="large")
    with left:
        st.markdown("**USD cost of capital**")
        st.metric("Equity cost of capital · USD", rate(result["re_us"]))
        st.metric("WACC · USD", rate(result["wacc_us"]))
        st.caption(f"Pretax debt cost: {rate(result['rd_us'])} · After-tax: {rate(result['rd_us']*(1-result['tax']))}")
    with right:
        st.markdown(f"**{country} cost of capital**")
        st.metric(f"Equity cost of capital · {country}", rate(result["re_local"]))
        st.metric(f"WACC · {country}", rate(result["wacc_local"]))
        st.caption(f"Pretax debt cost: {rate(result['rd_local'])} · After-tax: {rate(result['rd_local']*(1-result['tax']))}")
    st.caption(f"Relevered beta: {result['beta_e']:.4f} · Debt financing: {st.session_state.debt_pp}% · ERP: {st.session_state.erp_pp:.2f}% · Identical project tax rate in both currencies.")

    # Keep all assignment-required comparisons and downloads without overwhelming the landing page.
    with st.expander("All-industry comparison and CSV exports", expanded=False):
        comparison=[]
        for _,row in industry_frame.iterrows():
            r=compute(float(row.beta_u), st.session_state.debt_pp/100, tax_pp/100,
                  st.session_state.erp_pp/100, us_yield_pp/100, local_yield_pp/100,
                  usd_spread_pp/100, local_spread_pp/100 if foreign_override else None)
            comparison.append({"Industry":row.industry,"Unlevered beta":r['beta_u'],
                  "Levered beta":r['beta_e'],"USD equity cost (%)":r['re_us']*100,
                  "Foreign equity cost (%)":r['re_local']*100,
                  "USD debt cost (%)":r['rd_us']*100,"Foreign debt cost (%)":r['rd_local']*100,
                  "USD WACC (%)":r['wacc_us']*100,"Foreign WACC (%)":r['wacc_local']*100})
        if comparison:
            comparison_df=pd.DataFrame(comparison)
            st.dataframe(comparison_df, hide_index=True, use_container_width=True)
            st.download_button("Download all-industry comparison CSV", comparison_df.to_csv(index=False).encode(),
                               "all_industry_cost_of_capital.csv", mime="text/csv")
        else:
            st.warning("Damodaran beta dataset unavailable, so all-industry comparison cannot be generated. Refresh the source data.")
        export={"Industry":industry,"Country":country,"Unlevered beta":beta_u,
                "Relevered beta":result['beta_e'],"ERP (%)":st.session_state.erp_pp,
                "Debt share (%)":st.session_state.debt_pp,"Project tax jurisdiction":"Foreign" if use_local_tax else "US",
                "Applied tax (%)":tax_pp,"US 10-year yield (%)":us_yield_pp,"Foreign 10-year yield (%)":local_yield_pp,
                "USD BBB spread (%)":usd_spread_pp,"Foreign BBB spread (%)":result['spread_local']*100,
                "Foreign spread override":foreign_override,"USD equity cost (%)":result['re_us']*100,
                "Foreign equity cost (%)":result['re_local']*100,"USD debt cost (%)":result['rd_us']*100,
                "Foreign debt cost (%)":result['rd_local']*100,"USD WACC (%)":result['wacc_us']*100,
                "Foreign WACC (%)":result['wacc_local']*100,"Fetch time UTC":info['fetched_at'],
                "Beta source":info['beta_status'],"Spread source":info['spread_status'],
                "Tax source":info['tax_status'],"Yield source":info['yield_status']}
        st.download_button("Download this project's inputs and results CSV",
                           pd.DataFrame([export]).to_csv(index=False).encode(),
                           "cost_of_capital_selected_project.csv", mime="text/csv", type="primary")

    with st.expander("Calculation methodology and data sources", expanded=False):
        st.markdown(r"""
**Professor Rebelo's lecture formulas (slides 578–585):**

- Debt-to-equity: \(D/E=d/(1-d)\).
- Relevered equity beta: \(\beta_E=\beta_U[1+(1-\tau)D/E]\).
- USD cost of equity: \(R_e=R_{f,USD}+\beta_E ERP\).
- USD borrowing cost: \(R_d=R_{f,USD}+s_{USD}\).
- Gross currency factor: \(F=(1+R_{f,foreign})/(1+R_{f,USD})\).
- Foreign equity: \(R_e^*=(1+R_e)F-1\).
- Foreign spread: \(s^*=s_{USD}F\) unless overridden.
- Foreign debt: \(R_d^*=R_{f,foreign}+s^*\).
- Both WACCs: \(WACC=(1-d)R_e+dR_d(1-\tau)\).

**Assumptions:** U.S. market is the systematic-risk benchmark; government yields proxy for risk-free rates; equal real rates and relative PPP guide currency conversion. No additional country equity premium is included. The same debt share, ERP and project tax rate apply in both currencies. Source websites can block scraping; dated fallbacks and manual overrides are disclosed rather than presented as current observations.
""")
        st.markdown(f"- [Damodaran beta source]({BETA_URL}) — {info['beta_status']}")
        st.markdown(f"- [Damodaran borrowing-spread source]({SPREAD_URL}) — {info['spread_status']}")
        st.markdown(f"- [OECD tax source]({TAX_URL}) — {info['tax_status']}")
        st.markdown(f"- [Trading Economics yields]({TE_URL}) — {info['yield_status']}")
        st.caption(f"Retrieval attempt: {info['fetched_at']}")
        for err in info['errors']:
            st.warning(err)
    with st.expander("Independent test: Colombia lecture example"):
        check=compute(.80,.30,.35,.065,.04,.12,.015)
        st.write({"Relevered beta":round(check['beta_e'],6),
          "USD equity (%)":round(check['re_us']*100,4),"USD WACC (%)":round(check['wacc_us']*100,4),
          "Colombia equity (%)":round(check['re_local']*100,4),
          "Colombia WACC (%)":round(check['wacc_local']*100,4)})

if __name__ == "__main__":
    main()
