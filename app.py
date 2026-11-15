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
    st.title("🌐 International Cost of Capital Explorer")
    st.caption("Project cost of equity, cost of debt, and WACC in USD and a chosen currency · Professor Sergio Rebelo's 2026 lecture methodology")
    if "refresh_counter" not in st.session_state:
        st.session_state.refresh_counter = 0
    with st.sidebar:
        st.header("Project assumptions")
        country = st.selectbox("Target country", country_options(), index=country_options().index(US_NAME))
        with st.form("core_assumptions"):
            erp_pp = st.number_input("Equity risk premium (%)", min_value=0.0,max_value=50.0,value=6.5,step=0.25)
            debt_pp = st.slider("Target debt financing (%)",min_value=0,max_value=95,value=30,step=1)
            tax_jurisdiction = st.radio("Project tax rate used in BOTH currencies", ["U.S. corporate tax rate","Selected-country tax rate"], horizontal=False)
            st.form_submit_button("Apply assumptions")
        if st.button("↻ Refresh data from source websites", use_container_width=True):
            st.session_state.refresh_counter += 1
            st.rerun()
        st.caption("Refresh attempts all four source downloads. Some websites block automated retrieval; dated snapshots and manual overrides are indicated below.")
    with st.spinner("Checking source websites (first load can take a little while)…"):
        info = source_bundle(st.session_state.refresh_counter)
    industry_frame = info["beta"]
    industries = sorted(set(industry_frame["industry"].tolist()+[DEFAULT_INDUSTRY]))
    industry = st.selectbox("U.S. industry (Damodaran)", industries,
                            index=industries.index(DEFAULT_INDUSTRY))
    exact = industry_frame[industry_frame["industry"].str.casefold()==industry.casefold()]
    # Allow small source-label differences without guessing unrelated industries.
    if exact.empty and industry == DEFAULT_INDUSTRY and not industry_frame.empty:
        exact = industry_frame[industry_frame["industry"].str.contains(r"software.*(system|application)", case=False, regex=True)]
    beta_from_source = float(exact.iloc[0]["beta_u"]) if not exact.empty else None
    tax_us = info["tax"].get(US_NAME)
    tax_local = info["tax"].get(country)
    rf_us = info["yields"].get(US_NAME)
    rf_local = info["yields"].get(country)
    spread = info["spread"]

    st.subheader("Market data and overrides")
    st.caption("All rates below are in percent per annum. Checked overrides replace downloaded or dated source values; no hidden estimates are used.")
    with st.container(border=True):
        c1,c2,c3 = st.columns(3)
        with c1:
            st.markdown("**Operating risk and credit**")
            auto_beta = beta_from_source
            if auto_beta is None:
                st.warning("No downloaded beta for this industry. Supply a manual unlevered beta; illustrative default 0.80 comes from the lecture's Colombia example, NOT from Software industry data.")
            beta_override = st.checkbox("Override unlevered beta", value=auto_beta is None)
            beta_u = st.number_input("Unlevered beta βᵤ",min_value=0.0,max_value=10.0,
                    value=float(auto_beta if auto_beta is not None else 0.80),step=0.05,
                    disabled=not beta_override, key="beta_u_manual") if beta_override else auto_beta
            if beta_u is None:
                st.error("Unlevered beta is required")
                st.stop()
            if spread is None:
                st.warning("Live BBB spread unavailable. 1.50% below is the lecture's illustrative Colombia spread, NOT a current Damodaran quote.")
            spread_override = st.checkbox("Override USD BBB credit spread", value=spread is None)
            usd_spread_pp = st.number_input("USD BBB spread (%)",min_value=0.0,max_value=40.0,
                     value=float(spread if spread is not None else 1.50),step=0.05,
                     disabled=not spread_override) if spread_override else spread
        with c2:
            st.markdown("**U.S. market inputs**")
            override_us_yield = st.checkbox("Override U.S. 10-year yield")
            us_yield_pp = st.number_input("U.S. 10-year government yield (%)",min_value=-10.0,max_value=60.0,
                     value=float(rf_us if rf_us is not None else 0),step=0.1,
                     disabled=not override_us_yield)
            if rf_us is None and not override_us_yield:
                st.error("U.S. yield unavailable. Enable manual override.")
                st.stop()
            override_us_tax = st.checkbox("Override U.S. corporate tax rate")
            us_tax_pp = st.number_input("U.S. corporate tax (%)",min_value=0.0,max_value=65.0,
                     value=float(tax_us if tax_us is not None else 0),step=0.1,
                     disabled=not override_us_tax)
            if tax_us is None and not override_us_tax:
                st.error("U.S. corporate tax unavailable. Enable manual override.")
                st.stop()
        with c3:
            st.markdown(f"**{country} market inputs**")
            override_local_yield = st.checkbox("Override selected-country 10-year yield")
            local_yield_pp = st.number_input("Selected-country 10-year yield (%)",min_value=-10.0,max_value=100.0,
                    value=float(rf_local if rf_local is not None else 0),step=0.1,
                    disabled=not override_local_yield)
            if rf_local is None and not override_local_yield:
                st.error(f"No government yield available for {country}. Enable its manual override.")
                st.stop()
            override_local_tax = st.checkbox("Override selected-country corporate tax")
            local_tax_pp = st.number_input("Selected-country corporate tax (%)",min_value=0.0,max_value=65.0,
                     value=float(tax_local if tax_local is not None else 0),step=0.1,
                     disabled=not override_local_tax)
            if tax_local is None and not override_local_tax and tax_jurisdiction == "Selected-country tax rate":
                st.error("Selected-country tax data unavailable. Enable manual tax override.")
                st.stop()
            local_spread_override = st.checkbox("Override foreign-currency BBB spread", value=False)
            local_spread_pp = st.number_input("Foreign-currency BBB spread (%)",min_value=0.0,max_value=40.0,
                     value=float(usd_spread_pp*(1+local_yield_pp/100)/(1+us_yield_pp/100)),
                     step=0.05,disabled=not local_spread_override)
    project_tax_pp = us_tax_pp if tax_jurisdiction.startswith("U.S.") else local_tax_pp
    result = compute(beta_u,debt_pp/100,project_tax_pp/100,erp_pp/100,us_yield_pp/100,
                     local_yield_pp/100,usd_spread_pp/100,
                     local_spread_pp/100 if local_spread_override else None)
    st.caption(f"**Applied tax rate: {project_tax_pp:.2f}%** ({tax_jurisdiction}). Both currency scenarios use this exact rate. "
               f"Relevered equity beta: **{result['beta_e']:.4f}** · Target debt/equity: **{result['d_e']:.4f}** · "
               f"Foreign spread: **{rate(result['spread_local'])}** ({'manual' if local_spread_override else 'auto-converted'}).")

    st.subheader("Cost of capital · side-by-side")
    usd,local=st.columns(2)
    for col,side,name in [(usd,"us","USD"),(local,"local",f"{country} currency")]:
        with col:
            with st.container(border=True):
                st.markdown(f"#### {name}")
                st.metric("WACC",rate(result[f"wacc_{side}"]))
                a,b=st.columns(2)
                a.metric("Cost of equity", rate(result[f"re_{side}"]))
                b.metric("Pretax cost of debt",rate(result[f"rd_{side}"]))
                st.caption(f"After-tax cost of debt: {rate(result[f'rd_{side}']*(1-result['tax']))}")

    graph_rows=[]
    for label,key in [("Cost of equity","re"),("Pretax debt cost","rd"),("WACC","wacc")]:
        graph_rows.append({"Component":label,"Currency":"USD","Rate (%)":100*result[f"{key}_us"]})
        graph_rows.append({"Component":label,"Currency":country,"Rate (%)":100*result[f"{key}_local"]})
    chart=alt.Chart(pd.DataFrame(graph_rows)).mark_bar().encode(
        x=alt.X("Component:N",title=None),y=alt.Y("Rate (%):Q",title="Annualized percent"),
        color="Currency:N",xOffset="Currency:N",tooltip=["Component","Currency",alt.Tooltip("Rate (%):Q",format=".4f")])
    st.altair_chart(chart, use_container_width=True)

    st.subheader("All-industry comparison")
    if industry_frame.empty:
        st.warning("Cannot construct an actual all-industry comparison until Damodaran's beta workbook downloads successfully. Use Refresh; the selected-project manual beta still works.")
    else:
        comparison=[]
        for _,row in industry_frame.iterrows():
            r=compute(float(row.beta_u),debt_pp/100,project_tax_pp/100,erp_pp/100,
                      us_yield_pp/100,local_yield_pp/100,usd_spread_pp/100,
                      local_spread_pp/100 if local_spread_override else None)
            comparison.append({"Industry":row.industry,"Unlevered beta":r["beta_u"],
                "Levered beta":r["beta_e"], "USD equity cost (%)":r["re_us"]*100,
                "Foreign equity cost (%)":r["re_local"]*100,
                "USD debt cost (%)":r["rd_us"]*100,"Foreign debt cost (%)":r["rd_local"]*100,
                "USD WACC (%)":r["wacc_us"]*100,"Foreign WACC (%)":r["wacc_local"]*100})
        comparison_df=pd.DataFrame(comparison)
        st.dataframe(comparison_df,hide_index=True,use_container_width=True)
        st.download_button("Download all-industry comparison (CSV)",comparison_df.to_csv(index=False).encode("utf-8"),
                           "international_cost_of_capital_all_industries.csv",mime="text/csv")

    st.subheader("Download selected-project calculations")
    export={"Timestamp (UTC)":info["fetched_at"],"Industry":industry,"Country":country,
       "Beta source":info["beta_status"] if not beta_override else "Manual beta (or illustrative starting value)",
       "Unlevered beta":beta_u,"Levered beta":result["beta_e"],
       "ERP (%)":erp_pp,"Debt share (%)":debt_pp,"Tax source selection":tax_jurisdiction,
       "Applied tax (%)":project_tax_pp,"U.S. yield (%)":us_yield_pp,"Foreign yield (%)":local_yield_pp,
       "USD BBB spread (%)":usd_spread_pp,"Foreign BBB spread (%)":result["spread_local"]*100,
       "Foreign spread manual":local_spread_override,"USD equity (%)":result["re_us"]*100,
       "Foreign equity (%)":result["re_local"]*100,"USD debt (%)":result["rd_us"]*100,
       "Foreign debt (%)":result["rd_local"]*100,"USD WACC (%)":result["wacc_us"]*100,
       "Foreign WACC (%)":result["wacc_local"]*100,"Beta URL":BETA_URL,
       "Credit URL":SPREAD_URL,"Tax URL":TAX_URL,"Government yield URL":TE_URL,
       "Source beta status":info["beta_status"],"Source BBB status":info["spread_status"],
       "Source tax status":info["tax_status"],"Source yields status":info["yield_status"]}
    st.download_button("Download selected project (CSV)",pd.DataFrame([export]).to_csv(index=False).encode(),
                        "international_cost_of_capital_project.csv",mime="text/csv",type="primary")

    with st.expander("Formulas, assumptions, sources and retrieval notes",expanded=True):
        st.markdown(r"""
**Lecture equations (slides 579–585):**

- Debt/equity: \(D/E=d/(1-d)\).
- Relevered equity beta: \(\beta_E=\beta_U[1+(1-\tau)D/E]\).
- USD equity cost: \(R_e=R_{f,USD}+\beta_E\times ERP\).
- USD pretax debt cost: \(R_d=R_{f,USD}+s\), where \(s\) is the Baa2/BBB spread.
- Gross currency adjustment factor: \(F=(1+R_{f,foreign})/(1+R_{f,USD})\).
- Foreign equity cost: \(R_e^*=(1+R_e)F-1\).
- Auto-converted foreign borrowing spread: \(s^*=sF\), unless manually overridden.
- Foreign pretax debt cost: \(R_d^*=R_{f,foreign}+s^*\).
- In either currency: \(WACC=(1-d)R_e+dR_d(1-\tau)\).

**Interpretation:** Government bond yields are proxies for risk-free rates, rather than perfectly risk-free investments. Currency conversion assumes comparable real rates and relative purchasing power parity. No separate country equity risk premium is added. The same target leverage, ERP, and chosen tax rate apply in both scenarios. A selected-country currency does not mean spot FX conversion of rates.

**Data limitations:** Live scraping of Trading Economics and OECD may be blocked or their markup may change. Dates and source status are disclosed; dated lecture snapshots are not real-time quotes. For investment use, independently confirm quoted inputs. The manual override gives you control when a series cannot be retrieved. No API key is required by this app; Trading Economics may restrict scraping.
""")
        st.markdown(f"- [Damodaran industry unlevered betas]({BETA_URL}) — {info['beta_status']}")
        st.markdown(f"- [Damodaran BBB spread]({SPREAD_URL}) — {info['spread_status']}")
        st.markdown(f"- [OECD corporate income tax rates]({TAX_URL}) — {info['tax_status']}")
        st.markdown(f"- [Trading Economics government yields]({TE_URL}) — {info['yield_status']}")
        st.caption(f"Last attempted download: {info['fetched_at']}")
        for err in info["errors"]:
            st.warning(err)
    with st.expander("Built-in lecture calculation check (Colombia, slides 578–585)"):
        check=compute(.8,.30,.35,.065,.04,.12,.015)
        st.write({"Relevered beta":round(check["beta_e"],6),
                  "USD equity (%)":round(check["re_us"]*100,4),
                  "USD pretax debt (%)":round(check["rd_us"]*100,4),
                  "USD WACC (%)":round(check["wacc_us"]*100,4),
                  "COP equity (%)":round(check["re_local"]*100,4),
                  "COP pretax debt (%)":round(check["rd_local"]*100,4),
                  "COP WACC (%)":round(check["wacc_local"]*100,4)})
        st.caption("These are the slides' illustrative inputs, not current Colombia market observations.")

if __name__ == "__main__":
    main()
