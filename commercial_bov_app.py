"""
Commercial Property Advisory BOV
Estupinan Group | First Service Realty by ERA
Standard commercial property valuation for listings, acquisitions, and advisory assignments.
"""
import streamlit as st
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import io, re
from datetime import datetime, date, timedelta
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
from pathlib import Path
from PIL import Image as PILImage
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib.colors import HexColor, black, white
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT, TA_JUSTIFY
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer as RLSpacer, Table, TableStyle, PageBreak, Image, HRFlowable, KeepTogether)
from sqlalchemy import create_engine, text

def _check_password():
    def _submitted():
        if st.session_state.get("_pw_input") == st.secrets.get("APP_PASSWORD", ""):
            st.session_state["_authenticated"] = True
        else:
            st.session_state["_authenticated"] = False
            st.session_state["_pw_failed"] = True

    if st.session_state.get("_authenticated"):
        return True

    st.markdown("## Estupinan Group — Secure Access")
    st.markdown("**Commercial Property Advisory BOV** — Enter your password to continue.")
    st.text_input("Password", type="password", key="_pw_input", on_change=_submitted)
    if st.session_state.get("_pw_failed"):
        st.error("Incorrect password. Please try again.")
    st.stop()

_check_password()

BRAND_NAME = "Estupinan Group"
BRAND_AGENT = "Madelin Estupinan, Realtor\u00ae | Principal/Founder"
BRAND_CONTACT = "(786) 514-8046"
BRAND_FULL = f"{BRAND_AGENT}, {BRAND_NAME} | {BRAND_CONTACT}"
LICENSE_AGENT = "SL3672701"
LICENSE_BROKER = "BK3672701"
BROKER_FIRM = "Estupinan Group | First Service Realty by ERA"

BG_DARK="#0E1117"; BG_CARD="#161B22"; BORDER="#30363D"; NAVY="#0B1D3A"
GOLD="#D4A843"; RED="#E5534B"; RED_BG="#2D1B1B"; GREEN="#3FB950"
GREEN_BG="#1B2D1B"; TEXT_PRIMARY="#E6EDF3"; TEXT_SECONDARY="#8B949E"; TEXT_MUTED="#6E7681"

PROPERTY_TYPES = ["Multifamily","Retail Strip Center","Retail — Single Tenant NNN","Office — Single Tenant","Office — Multi-Tenant","Industrial/Warehouse","Industrial — Flex Space","Mixed-Use","Hospitality","Self-Storage","Medical Office","Land — Commercial","Land — Industrial","Land — Multifamily Development","Special Purpose"]
CONDITIONS = ["Excellent","Good","Fair","Poor","Requires Demolition"]
COND_RANK = {"Excellent":4,"Good":3,"Fair":2,"Poor":1,"Requires Demolition":0}
SALE_TYPES = ["Arms-length","REO/Distressed","Short Sale","Related Party","Active Listing"]
FLOOD_ZONES = ["X","AE","AH","VE","A","V","D","Other"]
MKT_TIMES = ["30-60 days","60-90 days","90-120 days","120-180 days","180+ days"]

def fmt_d(v):
    if v is None or v==0: return "$0"
    if v<0: return f"-${abs(v):,.0f}"
    return f"${v:,.0f}"
def fmt_d2(v):
    if v is None or v==0: return "$0.00"
    return f"${v:,.2f}"
def sl(text):
    return f'<div class="section-label">{text}</div>'
def sanitize_pdf_text(value):
    if value is None:
        return ""
    return str(value).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("\n", "<br/>")

st.set_page_config(page_title="Commercial Property Advisory BOV", page_icon="🏢", layout="wide", initial_sidebar_state="collapsed")
local_db_path = Path(__file__).resolve().with_name("bov_history_local.db") if "__file__" in globals() else Path.cwd() / "bov_history_local.db"
db_conn = None
local_engine = None
cloud_db_connected = False
try:
    db_conn = st.connection("postgresql", type="sql")
    with db_conn.session as session:
        session.execute(text("""
            CREATE TABLE IF NOT EXISTS commercial_bov_history (
                id BIGSERIAL PRIMARY KEY,
                address TEXT,
                ptype TEXT,
                dt TEXT,
                rec_value NUMERIC(15,2),
                sale_price NUMERIC(15,2),
                dom INTEGER,
                generated_by TEXT,
                created TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """))
        session.execute(text("ALTER TABLE commercial_bov_history ADD COLUMN IF NOT EXISTS generated_by TEXT"))
        session.commit()
    cloud_db_connected = True
except Exception:
    local_engine = create_engine(f"sqlite:///{local_db_path}")
    try:
        with local_engine.begin() as connection:
            connection.execute(text("""
                CREATE TABLE IF NOT EXISTS commercial_bov_history (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    address TEXT,
                    ptype TEXT,
                    dt TEXT,
                    rec_value NUMERIC,
                    sale_price NUMERIC,
                    dom INTEGER,
                    generated_by TEXT,
                    created TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """))
    except Exception:
        local_engine = None
if "inst_mode" not in st.session_state: st.session_state.inst_mode = False
if "analyst_name" not in st.session_state: st.session_state.analyst_name = "Madelin Estupinan"
st.sidebar.markdown(sl("Advisory Team"), unsafe_allow_html=True)
st.sidebar.selectbox("Lead Analyst", ["Madelin Estupinan", "Eric Yi", "Eddie San Roman"], key="analyst_name")
st.sidebar.markdown('---')
st.sidebar.toggle("Institutional Mode (High-Value Asset)", key="inst_mode")
st.sidebar.caption("Minimizes franchise branding and applies Private Equity slate-navy styling for institutional pitches.")
st.sidebar.markdown('---')
if cloud_db_connected:
    st.sidebar.success("Cloud Sync Active")
else:
    st.sidebar.warning("Running in Offline Mode — Local Storage Only")
st.sidebar.info("Commercial Advisory BOV — Build 1.0 | April 2026")
st.sidebar.divider()
if st.sidebar.button("Clear Current Analysis", use_container_width=True):
    preserved_state = {
        "inst_mode": st.session_state.get("inst_mode", False),
        "analyst_name": st.session_state.get("analyst_name", "Madelin Estupinan"),
    }
    st.session_state.clear()
    for key, value in preserved_state.items():
        st.session_state[key] = value
    st.rerun()
inst_mode = st.session_state.get("inst_mode", False)
brand_accent = "#94A3B8" if inst_mode else "#D4A843"
brand_subline = "Estupinan Group | Commercial Advisory | Institutional Mode" if inst_mode else BRAND_FULL
st.markdown(""" 
<style>
    .stApp { background-color: #0E1117; }
    .block-container { padding-top: 1.5rem; max-width: 1300px; }
    #MainMenu, footer, header { visibility: hidden; }
    div[data-testid="stDecoration"] { display: none; }
    .brand-bar { border-bottom: 1px solid #30363D; padding-bottom: 1rem; margin-bottom: 0.8rem; }
    .brand-bar h1 { color: BRAND_ACCENT; font-size: 1.5rem; font-weight: 700; margin: 0; }
    .brand-bar .sub { color: #8B949E; font-size: 0.78rem; margin-top: 0.15rem; }
    .section-label { color: #8B949E; font-size: 0.7rem; font-weight: 600; text-transform: uppercase; letter-spacing: 0.08em; padding-bottom: 0.4rem; margin-top: 0.8rem; margin-bottom: 0.3rem; border-bottom: 1px solid #30363D; }
    .mc { background: #161B22; border: 1px solid #30363D; border-radius: 8px; padding: 0.7rem 1rem; margin-bottom: 0.4rem; }
    .mc.alert { border-color: #E5534B; background: #2D1B1B; } .mc.ok { border-color: #3FB950; background: #1B2D1B; }
    .mc.gold { border-color: #8B7335; } .mc.blue { border-color: #58A6FF; }
    .ml { font-size: 0.65rem; color: #6E7681; text-transform: uppercase; letter-spacing: 0.06em; margin-bottom: 0.1rem; }
    .mv { font-size: 1.4rem; font-weight: 700; color: #E6EDF3; line-height: 1.2; }
    .mv.red { color: #E5534B; } .mv.green { color: #3FB950; } .mv.gold { color: #D4A843; } .mv.blue { color: #58A6FF; }
    .co { background: #161B22; border: 1px solid #30363D; border-left: 3px solid #D4A843; border-radius: 4px; padding: 0.6rem 0.8rem; font-size: 0.78rem; color: #8B949E; margin: 0.5rem 0; }
    .co.warn { border-left-color: #E5534B; } .co.info { border-left-color: #58A6FF; }
    .co strong { color: #E6EDF3; }
    .rr { display: flex; justify-content: space-between; padding: 0.3rem 0; border-bottom: 1px solid #30363D; font-size: 0.82rem; }
    .rr:last-child { border-bottom: none; }
    .rr .l { color: #8B949E; } .rr .v { color: #E6EDF3; font-weight: 600; } .rr .v.gold { color: #D4A843; }
    .dv { border-top: 1px solid #30363D; margin: 1rem 0; }
    .ft { text-align: center; color: #6E7681; font-size: 0.65rem; padding: 1rem 0 0.5rem 0; border-top: 1px solid #30363D; margin-top: 1.5rem; }
    label { color: #8B949E !important; font-size: 0.78rem !important; }
    div[data-testid="stVerticalBlock"] > div { gap: 0.2rem; }
</style>
""".replace("BRAND_ACCENT", brand_accent), unsafe_allow_html=True)
st.markdown(f'<div class="brand-bar"><h1>Commercial Property Advisory — Broker Opinion of Value</h1><div class="sub">{brand_subline}</div></div>', unsafe_allow_html=True)
if "nc" not in st.session_state: st.session_state.nc = 3
if "last_recon_val" not in st.session_state: st.session_state.last_recon_val = 0.0
if "weight_error" not in st.session_state: st.session_state.weight_error = False
if "conf_score" not in st.session_state: st.session_state.conf_score = "N/A"
if "cov" not in st.session_state: st.session_state.cov = 0.0
if "opex_ins" not in st.session_state: st.session_state.opex_ins = 0.0
if "effective_date" not in st.session_state: st.session_state.effective_date = date.today()
if "intended_use" not in st.session_state: st.session_state.intended_use = "Listing Price Guidance"
if "intended_user" not in st.session_state: st.session_state.intended_user = ""
if "oreo_days_remaining" not in st.session_state: st.session_state.oreo_days_remaining = None
if "oreo_deadline" not in st.session_state: st.session_state.oreo_deadline = None
if "oreo_months_remaining" not in st.session_state: st.session_state.oreo_months_remaining = None
if "oreo_months_held" not in st.session_state: st.session_state.oreo_months_held = None
if "dscr" not in st.session_state: st.session_state.dscr = None
if "annual_debt_service" not in st.session_state: st.session_state.annual_debt_service = None
if "loan_amount" not in st.session_state: st.session_state.loan_amount = None
if "loan_ltv" not in st.session_state: st.session_state.loan_ltv = None
if "loan_rate" not in st.session_state: st.session_state.loan_rate = None
if "loan_amort_years" not in st.session_state: st.session_state.loan_amort_years = None
if "irr_value" not in st.session_state: st.session_state.irr_value = None
if "ddr" not in st.session_state: st.session_state.ddr = 8.5
if "use_dcf" not in st.session_state: st.session_state.use_dcf = False
if "dscr_no_income" not in st.session_state: st.session_state.dscr_no_income = False
if "breakeven_occ" not in st.session_state: st.session_state.breakeven_occ = None
if "redistribution_occurred" not in st.session_state: st.session_state.redistribution_occurred = False
recon_val = st.session_state.last_recon_val
tab1,tab2,tab3,tab4,tab5 = st.tabs(["\u2460 Subject Property","\u2461 Comparable Sales","\u2462 Income Approach","\u2463 Cost Approach","\u2464 Reconciliation & PDF"])

with tab1:
    st.markdown(sl("Property Identification"), unsafe_allow_html=True)
    c1,c2 = st.columns(2)
    with c1: subj_address = st.text_input("Property Address", key="sa", placeholder="e.g. 4795 SW 8th Street, Miami FL")
    with c2: subj_folio = st.text_input("Folio Number", key="sf", placeholder="e.g. 01-4138-011-0010")
    if subj_folio and len(subj_folio.replace("-", "").replace(" ", "")) >= 10:
        lookup_c1, lookup_c2 = st.columns([1, 3])
        with lookup_c1:
            folio_clean = subj_folio.replace("-", "").replace(" ", "")
            pa_url = f"https://apps.miamidadepa.gov/propertysearch/#/?folio={folio_clean}"
            st.link_button("↗ Open Miami-Dade PA Record", url=pa_url, use_container_width=True)
    c1,c2,c3 = st.columns(3)
    with c1: subj_type = st.selectbox("Property Type", PROPERTY_TYPES, key="st2")
    with c2: subj_cond = st.selectbox("Condition", CONDITIONS, key="sc")
    with c3: subj_yr = st.number_input("Year Built", min_value=1900, max_value=date.today().year + 1, value=1985, key="sy")
    st.markdown(sl("Building & Site"), unsafe_allow_html=True)
    c1,c2,c3 = st.columns(3)
    with c1: subj_sf = st.number_input("Building Size (SF)", min_value=0, value=0, step=500, key="ssf")
    with c2: subj_lot = st.number_input("Lot Size (SF)", min_value=0, value=0, step=500, key="sl2")
    with c3:
        lot_ac = subj_lot/43560 if subj_lot>0 else 0
        st.markdown(f'<div class="mc"><div class="ml">Lot (Acres)</div><div class="mv" style="font-size:1.1rem;">{lot_ac:.3f}</div></div>', unsafe_allow_html=True)
    c1,c2,c3 = st.columns(3)
    with c1: subj_units = st.number_input("Units (if MF)", min_value=0, value=0, key="su")
    with c2: subj_stories = st.number_input("Stories", min_value=0, value=1, key="ss")
    with c3: subj_parking = st.number_input("Parking Spaces", min_value=0, value=0, key="sp")
    c1,c2,c3 = st.columns(3)
    with c1: subj_occ = st.number_input("Occupancy (%)", min_value=0.0, max_value=100.0, value=0.0, step=5.0, key="so")
    with c2: subj_zone = st.text_input("Zoning", key="sz", placeholder="e.g. T6-8-O")
    with c3: subj_flood = st.selectbox("Flood Zone", FLOOD_ZONES, key="sfl")
    st.markdown(sl("Use & Description"), unsafe_allow_html=True)
    c1,c2 = st.columns(2)
    with c1: subj_curuse = st.text_input("Current Use", key="scu")
    with c2: subj_hbu_imp = st.text_input("HBU (as improved)", key="shi")
    c1,c2 = st.columns(2)
    with c1: subj_hbu_vac = st.text_input("HBU (as vacant)", key="shv")
    with c2: pass
    subj_desc = st.text_area("Property Description", key="sd", height=80, placeholder="2-3 sentences for the BOV report...")
    st.markdown(sl("Property Visuals"), unsafe_allow_html=True)
    subj_photo = st.file_uploader("Upload Subject Property Photo (Primary)", type=['jpg', 'png', 'jpeg'], key="subj_photo_upload")
    if subj_photo is not None:
        st.image(subj_photo, width=400, caption="Subject Property")
    st.markdown(sl("Subject Transaction History"), unsafe_allow_html=True)
    c1,c2 = st.columns(2)
    with c1: last_sale_date = st.date_input("Last Arm's-Length Sale Date", value=None, min_value=date(1900, 1, 1), key="lsd")
    with c2: last_sale_price = st.number_input("Last Sale Price ($)", min_value=0.0, value=0.0, step=25000.0, key="lsp")
    st.caption("Prior sale history for market trend and value context.")
    st.markdown(sl("Ownership & Title"), unsafe_allow_html=True)
    c1,c2 = st.columns(2)
    with c1: owner_of_record = st.text_input("Owner of Record / Current Titleholder", key="owner_of_record", placeholder="e.g. First National Bank, as Trustee")
    with c2: title_status = st.selectbox("Title / Ownership Status", ["Fee Simple — Arm's-Length Owner", "Trust / Estate", "LLC / Corporate Owned", "Partnership Owned", "Tenants in Common", "Ground Lease", "Other"], key="title_status")
    extraordinary_assumptions = st.text_area("Extraordinary Assumptions & Hypothetical Conditions", key="extraordinary_assumptions", height=80, value=st.session_state.get("extraordinary_assumptions", "This analysis assumes clear and marketable title free of undisclosed encumbrances, liens, or restrictions not identified herein. This analysis assumes the property is in the condition observed and described in this report. This analysis assumes no undisclosed environmental conditions exist on the property. This analysis assumes all information provided by the client or obtained from public records is accurate and complete."))
    st.caption("Extraordinary assumptions are facts assumed to be true for purposes of this analysis but which may not be verified. Hypothetical conditions are contrary to known fact. Both must be disclosed per standard BOV practice.")
    st.markdown(sl("Known Issues"), unsafe_allow_html=True)
    c1,c2 = st.columns(2)
    with c1:
        has_env = st.toggle("Environmental Issues", key="he")
        env_desc = st.text_input("Describe", key="ed") if has_env else ""
    with c2:
        has_code = st.toggle("Code Violations", key="hc")
        code_desc = st.text_input("Describe", key="cd") if has_code else ""
    st.markdown(sl("Report Details"), unsafe_allow_html=True)
    c1,c2,c3 = st.columns(3)
    with c1: prep_for = st.text_input("Prepared For", key="pf", placeholder="e.g. Property Owner, Buyer, Attorney")
    with c2: rpt_date = st.date_input("Date of Opinion", value=date.today(), key="rd")
    with c3: effective_date = st.date_input("Effective Date of Value", value=st.session_state.get("effective_date", date.today()), key="effective_date", help="Date as of which the value opinion is valid; may differ from the report date in retrospective valuations.")
    st.markdown(sl("Intended Use and User"), unsafe_allow_html=True)
    c1,c2 = st.columns(2)
    intended_use_options = [
        "Listing Price Guidance",
        "Buyer Acquisition Analysis",
        "Portfolio Review / Asset Management",
        "Lease-Up / Stabilization Advisory",
        "1031 Exchange Identification",
        "Partnership / JV Valuation",
        "Estate / Probate Valuation",
        "Refinance Advisory",
        "Other",
    ]
    default_use = st.session_state.get("intended_use", "Listing Price Guidance")
    default_use_index = intended_use_options.index(default_use) if default_use in intended_use_options else 0
    with c1: intended_use = st.selectbox("Intended Use of this BOV", intended_use_options, index=default_use_index, key="intended_use")
    with c2: intended_user = st.text_input("Intended User(s)", key="intended_user", placeholder="e.g. Property Owner, Buyer Representative — not for third-party reliance")
    st.session_state.oreo_days_remaining = None
    st.session_state.oreo_deadline = None
    st.session_state.oreo_months_remaining = None
    st.session_state.oreo_months_held = None
    use_letter = st.toggle("Include Transmittal Letter", key="use_letter")
    recipient_name = ""
    recipient_title = ""
    engagement_narrative = "At your request, we have performed an analysis to determine the Broker Opinion of Value for the referenced property. Our analysis considers current market conditions, comparable sales trends, and the income-producing potential of the asset."
    if use_letter:
        recipient_name = st.text_input("Recipient Name", key="recipient_name", placeholder="e.g. Mr. Juan Sanchez")
        recipient_title = st.text_input("Recipient Title / Company", key="recipient_title", placeholder="e.g. Managing Member, Asset Manager, Counsel")
        engagement_narrative = st.text_area(
            "Engagement Narrative",
            key="engagement_narrative",
            height=110,
            value=engagement_narrative,
        )

with tab2:
    st.markdown(sl("Comparable Sales (3-8 comps)"), unsafe_allow_html=True)
    st.markdown('<div class="co info">Enter comp data from your research. Includes closed sales and active listings. System auto-calculates $/SF, weights, and flags outliers.</div>', unsafe_allow_html=True)
    def build_comp_template():
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Comps"
        headers = [
            "Address",
            "Sale Type",
            "Sale/List Date (YYYY-MM-DD)",
            "Sale/List Price ($)",
            "Building SF",
            "Lot SF",
            "Year Built",
            "Renovation Year (0=none)",
            "Condition",
            "Occupancy (%)",
            "Units (MF only)",
            "Cap Rate (%)",
            "Distance (miles)",
            "Adjustment (%)",
            "Adjustment Notes",
            "Data Source",
        ]
        notes = [
            "Full property address",
            "Closed Sale / Active Listing / Expired / Withdrawn",
            "Format: 2024-06-15",
            "Total sale or list price",
            "Gross building area in SF",
            "Land area in SF",
            "4-digit year e.g. 1985",
            "4-digit year 0 if none",
            "Excellent / Good / Average / Fair / Poor",
            "0 to 100",
            "0 if not multifamily",
            "Leave 0 if unknown",
            "Miles from subject",
            "Positive subject superior negative subject inferior",
            "Brief description of adjustment rationale",
            "e.g. CoStar Public Records MLS",
        ]
        for col_idx, header in enumerate(headers, start=1):
            ws.cell(row=1, column=col_idx, value=header)
            ws.cell(row=2, column=col_idx, value=notes[col_idx - 1])
        header_fill = PatternFill(fill_type="solid", fgColor="1E3A5F")
        note_fill = PatternFill(fill_type="solid", fgColor="D6E4F0")
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = Font(color="FFFFFF", bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for cell in ws[2]:
            cell.fill = note_fill
            cell.font = Font(color="1E3A5F", italic=True)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for column_cells in ws.columns:
            ws.column_dimensions[column_cells[0].column_letter].width = 20
        ws.row_dimensions[1].height = 28
        ws.row_dimensions[2].height = 32
        for row_idx in range(3, 11):
            ws.row_dimensions[row_idx].height = 18
        buffer = io.BytesIO()
        wb.save(buffer)
        buffer.seek(0)
        return buffer.getvalue()

    comp_dl_col, comp_upload_col, comp_msg_col = st.columns([1, 1, 2])
    with comp_dl_col:
        st.download_button(
            label="⬇ Download Comp Template",
            data=build_comp_template(),
            file_name="Estupinan_Comp_Template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            help="Download the Excel template, fill in your comps, then upload it below.",
        )
    with comp_upload_col:
        uploaded_comps = st.file_uploader("Upload comp template", type=["xlsx"], key="comp_upload", label_visibility="collapsed")
    with comp_msg_col:
        if uploaded_comps is not None:
            st.markdown('<div class="co info">✅ Comp template uploaded. Fields pre-populated below — review and adjust as needed.</div>', unsafe_allow_html=True)

    if uploaded_comps is not None:
        upload_signature = (uploaded_comps.name, len(uploaded_comps.getvalue()), hash(uploaded_comps.getvalue()))
        if st.session_state.get("_comp_upload_signature") != upload_signature:
            try:
                workbook = openpyxl.load_workbook(io.BytesIO(uploaded_comps.getvalue()), data_only=True)
                worksheet = workbook["Comps"] if "Comps" in workbook.sheetnames else workbook.active
                filled_idx = 0
                for row in worksheet.iter_rows(min_row=3, max_row=10, values_only=True):
                    if filled_idx > 7:
                        break
                    if not row or row[0] in (None, ""):
                        continue

                    raw_sale_type = str(row[1]).strip() if row[1] not in (None, "") else ""
                    sale_type_norm = raw_sale_type.lower()
                    if sale_type_norm == "active listing":
                        sale_type = "Active Listing"
                    elif sale_type_norm in {"expired", "withdrawn"}:
                        sale_type = "Active Listing"
                    elif sale_type_norm == "reo/distressed":
                        sale_type = "REO/Distressed"
                    elif raw_sale_type in SALE_TYPES:
                        sale_type = raw_sale_type
                    else:
                        sale_type = "Arms-length"

                    raw_date = row[2]
                    comp_date = date.today() - timedelta(days=90 * (filled_idx + 1))
                    if isinstance(raw_date, datetime):
                        comp_date = raw_date.date()
                    elif isinstance(raw_date, date):
                        comp_date = raw_date
                    elif raw_date not in (None, ""):
                        try:
                            comp_date = datetime.strptime(str(raw_date).strip(), "%Y-%m-%d").date()
                        except ValueError:
                            pass

                    raw_condition = str(row[8]).strip() if row[8] not in (None, "") else ""
                    if raw_condition in CONDITIONS:
                        comp_condition = raw_condition
                    elif raw_condition.lower() == "average":
                        comp_condition = "Good"
                    else:
                        comp_condition = "Good"

                    def _safe_float(value, default=0.0):
                        try:
                            return float(value)
                        except (TypeError, ValueError):
                            return default

                    def _safe_int(value, default=0):
                        try:
                            return int(float(value))
                        except (TypeError, ValueError):
                            return default

                    st.session_state[f"ca{filled_idx}"] = str(row[0]).strip()
                    st.session_state[f"cs{filled_idx}"] = sale_type
                    st.session_state[f"cd{filled_idx}"] = comp_date
                    st.session_state[f"cp{filled_idx}"] = max(_safe_int(row[3]), 0)
                    st.session_state[f"csf{filled_idx}"] = max(_safe_int(row[4]), 0)
                    st.session_state[f"cl{filled_idx}"] = max(_safe_int(row[5]), 0)
                    st.session_state[f"cy{filled_idx}"] = max(_safe_int(row[6], 1985), 1900)
                    st.session_state[f"crn{filled_idx}"] = max(_safe_int(row[7]), 0)
                    st.session_state[f"cc{filled_idx}"] = comp_condition
                    st.session_state[f"co{filled_idx}"] = min(max(_safe_float(row[9]), 0.0), 100.0)
                    st.session_state[f"cu{filled_idx}"] = max(_safe_int(row[10]), 0)
                    st.session_state[f"cr{filled_idx}"] = min(max(_safe_float(row[11]), 0.0), 20.0)
                    st.session_state[f"cdi{filled_idx}"] = min(max(_safe_float(row[12]), 0.0), 20.0)
                    st.session_state[f"caj{filled_idx}"] = min(max(_safe_float(row[13]), -50.0), 50.0)
                    st.session_state[f"cn{filled_idx}"] = "" if row[14] in (None, "") else str(row[14]).strip()
                    st.session_state[f"cds{filled_idx}"] = "" if row[15] in (None, "") else str(row[15]).strip()
                    filled_idx += 1

                st.session_state["nc"] = max(3, min(8, filled_idx))
                st.session_state["nci"] = max(3, min(8, filled_idx))
                st.session_state["_comp_upload_signature"] = upload_signature
                st.rerun()
            except Exception as exc:
                st.error(f"Unable to import comp template: {exc}")
    _,rc = st.columns([3,1])
    with rc: num_c = st.number_input("Number of Comps", min_value=3, max_value=8, value=st.session_state.nc, key="nci")
    st.session_state.nc = num_c
    manual_weight_override = False
    total_manual_weight = 0.0
    weight_error_message = ""
    comps = []
    for i in range(num_c):
        with st.expander(f"Comparable {i+1}", expanded=(i<3)):
            a1,a2 = st.columns(2)
            with a1: ca = st.text_input("Address", key=f"ca{i}")
            with a2: cst = st.selectbox("Sale Type", SALE_TYPES, key=f"cs{i}")
            a1,a2,a3 = st.columns(3)
            with a1: cd = st.date_input("Sale/List Date", value=date.today()-timedelta(days=90*(i+1)), min_value=date(1900, 1, 1), key=f"cd{i}")
            with a2: cp = st.number_input("Sale/List Price ($)", min_value=0, value=0, step=25000, key=f"cp{i}")
            with a3: csf = st.number_input("Building SF", min_value=0, value=0, step=500, key=f"csf{i}")
            a1,a2,a3 = st.columns(3)
            with a1: cl = st.number_input("Lot SF", min_value=0, value=0, step=500, key=f"cl{i}")
            with a2: cyr = st.number_input("Year Built", min_value=1900, max_value=date.today().year + 1, value=1985, key=f"cy{i}")
            with a3: cren = st.number_input("Renovation Year (0=none)", min_value=0, max_value=date.today().year + 1, value=0, key=f"crn{i}")
            a1,a2,a3 = st.columns(3)
            with a1: ccond = st.selectbox("Condition", CONDITIONS, key=f"cc{i}")
            with a2: cocc = st.number_input("Occupancy (%)", min_value=0.0, max_value=100.0, value=0.0, step=5.0, key=f"co{i}")
            with a3: cu = st.number_input("Units (MF)", min_value=0, value=0, key=f"cu{i}")
            a1,a2,a3 = st.columns(3)
            with a1: ccap = st.number_input("Cap Rate (%)", min_value=0.0, max_value=20.0, value=0.0, step=0.25, key=f"cr{i}")
            with a2: cdist = st.number_input("Distance (miles)", min_value=0.0, max_value=20.0, value=1.0, step=0.5, key=f"cdi{i}")
            with a3: cadj = st.number_input("Adjustment (%)", min_value=-50.0, max_value=50.0, value=0.0, step=1.0, key=f"caj{i}", help="+ subject superior, - subject inferior")
            cnotes = st.text_input("Adjustment Notes", key=f"cn{i}", placeholder="e.g. inferior location, superior condition")
            cphoto = st.file_uploader(f"Comp {i+1} Photo", type=['jpg', 'png', 'jpeg'], key=f"cp_photo_{i}")
            csource = st.text_input("Data Source", key=f"cds{i}", placeholder="e.g. CoStar, Public Records, Internal")
            psf = cp/csf if csf>0 else 0
            comps.append(dict(n=i+1, addr=ca, source=csource, photo=cphoto, dt=cd, price=cp, sf=csf, lot=cl, yr=cyr, ren=cren, cond=ccond, occ=cocc, units=cu, stype=cst, cap=ccap, dist=cdist, adj=cadj, notes=cnotes, psf=psf, apsf=psf*(1+cadj/100) if csf>0 else 0))
    use_land_sales = subj_type == "Land"
    use_unit_sales = subj_type == "Multifamily" and subj_sf == 0 and subj_units > 0
    sales_metric_label = "$/SF (Land)" if use_land_sales else ("$/Unit" if use_unit_sales else "$/SF")
    sales_metric_suffix = "/SF" if use_land_sales else ("/Unit" if use_unit_sales else "/SF")
    subject_basis = subj_lot if use_land_sales else (subj_units if use_unit_sales else subj_sf)
    vc = [
        c for c in comps
        if c["price"] > 0 and (
            (use_land_sales and c["lot"] > 0) or
            (use_unit_sales and c["units"] > 0) or
            ((not use_land_sales) and (not use_unit_sales) and c["sf"] > 0)
        )
    ]
    weighted_psf = 0; sales_comp_value = 0
    flagged_outliers = []
    if vc:
        st.markdown(sl("Analysis"), unsafe_allow_html=True)
        st.markdown(sl("Market Conditions Adjustment"), unsafe_allow_html=True)
        market_trend_rate = st.number_input("Annual Market Appreciation / Depreciation (%)", min_value=-25.0, max_value=25.0, value=0.0, step=0.25, help="Applied to each comparable based on time elapsed since the sale date; use negative values for declining markets.", key="market_trend_rate")
        manual_weight_override = st.toggle("Enable Manual Weighting Override", key="manual_weight_override")
        for c in vc:
            comp_basis = c["lot"] if use_land_sales else (c["units"] if use_unit_sales else c["sf"])
            c["psf"] = (c["price"]/comp_basis) if comp_basis>0 else 0
            months_elapsed = max((date.today() - c["dt"]).days / 30.44, 0.0)
            c["trend_factor"] = 1 + ((market_trend_rate / 100) * (months_elapsed / 12))
            c["trend_pct"] = (c["trend_factor"] - 1) * 100
            c["apsf"] = c["psf"] * c["trend_factor"] * (1+c["adj"]/100) if comp_basis>0 else 0
        apsfs = [c["apsf"] for c in vc]
        pm, pmed, plo, phi = np.mean(apsfs), np.median(apsfs), min(apsfs), max(apsfs)
        std_dev = np.std(apsfs)
        cov = (std_dev / pm) * 100 if pm > 0 else 0.0
        if cov < 10:
            conf_score, conf_color = "High", "green"
        elif cov <= 15:
            conf_score, conf_color = "Medium", "gold"
        else:
            conf_score, conf_color = "Low", "red"
        st.session_state.conf_score = conf_score
        st.session_state.cov = cov
        outliers = [c for c in vc if pmed > 0 and abs(c["apsf"]-pmed)/pmed>0.15]
        flagged_outliers = outliers.copy()
        today = date.today()
        for c in vc:
            mo = (today-c["dt"]).days/30.44
            c["wr"] = 1.0 if mo<=6 else (0.75 if mo<=12 else (0.50 if mo<=18 else 0.25))
            d=c["dist"]; c["wp"] = 1.0 if d<=0.5 else (0.80 if d<=1 else (0.60 if d<=3 else 0.40))
            sr=COND_RANK.get(subj_cond,2); cr=COND_RANK.get(c["cond"],2); diff=abs(sr-cr)
            c["wc"] = 1.0 if diff==0 else (0.80 if diff==1 else 0.60)
            c["wraw"] = c["wr"]*c["wp"]*c["wc"]
        tw = sum(c["wraw"] for c in vc)
        for c in vc: c["wn"] = c["wraw"]/tw if tw>0 else 1/len(vc)
        if manual_weight_override:
            st.markdown('<div class="co info">Manual weighting override is active. Enter broker-assigned comp weights below.</div>', unsafe_allow_html=True)
            mw_cols = st.columns(min(3, len(vc)))
            for idx, c in enumerate(vc):
                with mw_cols[idx % len(mw_cols)]:
                    manual_weight = st.number_input(f"Comp {c['n']} Weight (%)", min_value=0.0, max_value=100.0, value=0.0, step=1.0, key=f"mw_{c['n']}")
                c["wn"] = manual_weight / 100
            total_manual_weight = sum(c["wn"] for c in vc) * 100
        if not manual_weight_override:
            active_listings = [c for c in vc if c["stype"] == "Active Listing"]
            closed_sales = [c for c in vc if c["stype"] != "Active Listing"]
            redistribution_occurred = False
            if active_listings and closed_sales:
                excess_weight = 0.0
                for c in active_listings:
                    if c["wn"] > 0.10:
                        excess_weight += c["wn"] - 0.10
                        redistribution_occurred = True
                        c["wn"] = 0.10
                closed_weight_total = sum(c["wn"] for c in closed_sales)
                if excess_weight > 0 and closed_weight_total > 0:
                    for c in closed_sales:
                        c["wn"] += excess_weight * (c["wn"] / closed_weight_total)
            st.session_state.redistribution_occurred = redistribution_occurred
        else:
            active_listings = [c for c in vc if c["stype"] == "Active Listing"]
            closed_sales = [c for c in vc if c["stype"] != "Active Listing"]
            st.session_state.redistribution_occurred = False
            if active_listings:
                st.markdown('<div class="co warn"><strong>Manual Override Active:</strong> Active listings are present in your comp set. In manual mode, broker-assigned weights are applied without the institutional 10% cap. You accept full responsibility for the weighting of active listings in this BOV.</div>', unsafe_allow_html=True)
        weighted_psf = sum(c["apsf"]*c["wn"] for c in vc)
        sales_comp_value = weighted_psf*subject_basis if subject_basis>0 else 0
        m1,m2,m3,m4,m5 = st.columns(5)
        with m1: st.markdown(f'<div class="mc"><div class="ml">Mean {sales_metric_label}</div><div class="mv" style="font-size:1.1rem;">{fmt_d2(pm)}</div></div>', unsafe_allow_html=True)
        with m2: st.markdown(f'<div class="mc"><div class="ml">Median {sales_metric_label}</div><div class="mv" style="font-size:1.1rem;">{fmt_d2(pmed)}</div></div>', unsafe_allow_html=True)
        with m3: st.markdown(f'<div class="mc"><div class="ml">Range</div><div class="mv" style="font-size:0.95rem;">{fmt_d2(plo)} - {fmt_d2(phi)}</div></div>', unsafe_allow_html=True)
        with m4: st.markdown(f'<div class="mc gold"><div class="ml">Wtd Adj {sales_metric_label}</div><div class="mv gold">{fmt_d2(weighted_psf)}</div></div>', unsafe_allow_html=True)
        with m5: st.markdown(f'<div class="mc"><div class="ml">Statistical Confidence</div><div class="mv {conf_color}" style="font-size:1.1rem;">{conf_score}</div><div class="rr"><span class="l">COV</span><span class="v">{cov:.1f}%</span></div></div>', unsafe_allow_html=True)
        if subject_basis>0:
            basis_label = "Lot SF" if use_land_sales else ("Units" if use_unit_sales else "SF")
            st.markdown(f'<div class="mc ok"><div class="ml">Sales Comparison Value ({subject_basis:,} {basis_label} x {fmt_d2(weighted_psf)}{sales_metric_suffix})</div><div class="mv green">{fmt_d(sales_comp_value)}</div></div>', unsafe_allow_html=True)
        if manual_weight_override and abs(total_manual_weight-100.0) > 0.01:
            weight_error_message = f"CRITICAL: Manual weights sum to {total_manual_weight:.1f}%. They must sum to exactly 100% to generate a report."
            st.error(weight_error_message)
            st.session_state.weight_error = True
        else:
            st.session_state.weight_error = False
        if flagged_outliers:
            ns=", ".join([f"Comp {c['n']}" for c in flagged_outliers])
            st.markdown(f'<div class="co warn"><strong>Outlier flagged:</strong> {ns} - adjusted {sales_metric_label} is more than 15% from the adjusted median. Consider refining the comp set or market adjustments.</div>', unsafe_allow_html=True)
        outliers = flagged_outliers
        st.markdown(sl("Weighting Detail"), unsafe_allow_html=True)
        for c in vc:
            fl=' <span style="color:#E5534B;">OUTLIER</span>' if c in outliers else ""
            act = ' <span style="color:#58A6FF;">ACTIVE</span>' if c["stype"]=="Active Listing" else ""
            card_class = "mc blue" if manual_weight_override else "mc"
            st.markdown(f'<div class="{card_class}" style="padding:0.4rem 0.8rem;"><div style="display:flex;justify-content:space-between;font-size:0.8rem;"><span style="color:#D4A843;font-weight:700;">Comp {c["n"]}{act}</span><span style="color:#8B949E;">{c["addr"] or chr(8212)}{fl}</span></div><div class="rr"><span class="l">{sales_metric_label}: {fmt_d2(c["psf"])} | Trend ({c["trend_pct"]:+.1f}%): stacked before Adj ({c["adj"]:+.0f}%) = {fmt_d2(c["apsf"])}</span><span class="v gold">Wt: {c["wn"]:.1%}</span></div></div>', unsafe_allow_html=True)
        if active_listings and not manual_weight_override:
            if st.session_state.get("redistribution_occurred", False):
                st.markdown('<div class="co warn"><strong>Institutional treatment of active listings:</strong> Active listings are ceiling indicators only, not confirmed market evidence. One or more active listing weights exceeded 10% and were capped and redistributed to closed sales per institutional BOV methodology.</div>', unsafe_allow_html=True)
            else:
                st.markdown('<div class="co warn"><strong>Active listings present:</strong> Active listings are ceiling indicators only, not confirmed market evidence. Weights are monitored and limited to a maximum of 10% each per institutional BOV methodology.</div>', unsafe_allow_html=True)
    else:
        st.session_state.weight_error = False
        st.session_state.conf_score = "N/A"
        st.session_state.cov = 0.0
        empty_basis = "lot SF" if use_land_sales else ("units" if use_unit_sales else "building SF")
        st.markdown(f'<div class="co">Enter at least one comp with price and valid {empty_basis}.</div>', unsafe_allow_html=True)

with tab3:
    income_value=0; stab_value=0; dcf_value=0; noi=0; egi=0; total_opex=0; inc_cap=0
    gpr=0; reimb=0; vac_rate=5.0; opex_tax=0; opex_ins=0; opex_mgmt_pct=5.0; opex_maint=0
    opex_util=0; opex_res=0; opex_oth=0; opex_mgmt_d=0; opex_ratio=0; use_tax_reassessment=False; income_tax_line=0; mgmt_floor=0.0
    tenant_quality = "Investment Grade (S&P/Moody's Rated)"; risk_premium = 0.0; adjusted_cap = 0.0
    use_unit_mix = False; loss_to_lease = 0.0; rent_upside_pct = 0.0; total_market_mo = 0.0; market_total_annual = 0.0
    physical_vacancy_loss = 0.0; credit_loss_amount = 0.0; fixed_opex_base = 0.0; stab_egi = 0.0; stab_opex = 0.0
    dscr_value = None; annual_debt_service = None; loan_amount = None; loan_ltv = None; loan_rate = None; loan_amort_years = None; irr_value = None; dscr_no_income = False; breakeven_occ_pct = None
    tax_rate_pct = float(st.session_state.get("tax_rate_pct", 2.10)); ddr = st.session_state.get("ddr", 8.5); use_dcf = False
    unit_mix_summary = pd.DataFrame(columns=["Unit Type", "Count", "Vacant Count", "Total SF", "Monthly Rent", "Market Rent"])
    if subj_type == "Land":
        st.markdown('<div class="co info">Income approach not applicable for land.</div>', unsafe_allow_html=True)
    else:
        use_unit_mix = st.toggle("Enable Detailed Unit Mix / Rent Roll", key="use_unit_mix")
        st.markdown(sl("Current Income"), unsafe_allow_html=True)
        if use_unit_mix:
            rr_dl, rr_up = st.columns([1, 2])
            with rr_dl:
                def build_rr_template():
                    df = pd.DataFrame(columns=["Unit Type", "Count", "Vacant Count", "Avg. SF", "Current Rent/Mo", "Market Rent/Mo"])
                    buf = io.BytesIO()
                    df.to_excel(buf, index=False, engine='openpyxl')
                    return buf.getvalue()
                st.download_button("⬇ Download Rent Roll Template", build_rr_template(), "Estupinan_Rent_Roll.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            with rr_up:
                uploaded_rr = st.file_uploader("Upload Rent Roll", type=["xlsx"], key="rr_upload", label_visibility="collapsed")

            if uploaded_rr is not None:
                try:
                    uploaded_df = pd.read_excel(uploaded_rr)
                    required_cols = ["Unit Type", "Count", "Vacant Count", "Avg. SF", "Current Rent/Mo", "Market Rent/Mo"]
                    missing_cols = [col for col in required_cols if col not in uploaded_df.columns]
                    if missing_cols:
                        raise ValueError(f"Missing columns: {', '.join(missing_cols)}")
                    unit_mix_seed = uploaded_df
                except Exception as e:
                    st.error(f"Invalid template format: {e}. Please download and use the provided Rent Roll Template.")
                    unit_mix_seed = pd.DataFrame([{"Unit Type": "1BR", "Count": 0, "Vacant Count": 0, "Avg. SF": 0, "Current Rent/Mo": 0.0, "Market Rent/Mo": 0.0}])
            else:
                unit_mix_seed = pd.DataFrame([
                    {"Unit Type": "1BR", "Count": 0, "Vacant Count": 0, "Avg. SF": 0, "Current Rent/Mo": 0.0, "Market Rent/Mo": 0.0},
                    {"Unit Type": "2BR", "Count": 0, "Vacant Count": 0, "Avg. SF": 0, "Current Rent/Mo": 0.0, "Market Rent/Mo": 0.0},
                ])
            unit_mix_df = st.data_editor(unit_mix_seed, num_rows="dynamic", use_container_width=True, key="unit_mix_editor")
            unit_mix_df = pd.DataFrame(unit_mix_df).fillna(0)
            for col in ["Count", "Vacant Count", "Avg. SF", "Current Rent/Mo", "Market Rent/Mo"]:
                unit_mix_df[col] = pd.to_numeric(unit_mix_df[col], errors="coerce").fillna(0)
            unit_mix_df["Vacant Count"] = unit_mix_df["Vacant Count"].clip(lower=0)
            unit_mix_df["Vacant Count"] = unit_mix_df[["Vacant Count", "Count"]].min(axis=1)
            total_current_mo = (unit_mix_df["Count"] * unit_mix_df["Current Rent/Mo"]).sum()
            total_market_mo = (unit_mix_df["Count"] * unit_mix_df["Market Rent/Mo"]).sum()
            total_rent_roll_sf = (unit_mix_df["Count"] * unit_mix_df["Avg. SF"]).sum()
            if subj_sf > 0 and abs(total_rent_roll_sf - subj_sf) > (subj_sf * 0.01):
                st.warning(f"Data Discrepancy: Rent Roll SF ({int(total_rent_roll_sf):,}) does not match Building SF ({subj_sf:,})")
            gpr = total_current_mo * 12
            physical_vacancy_loss = (unit_mix_df["Vacant Count"] * unit_mix_df["Market Rent/Mo"]).sum() * 12
            occupied_count = (unit_mix_df["Count"] - unit_mix_df["Vacant Count"]).clip(lower=0)
            loss_to_lease = (occupied_count * (unit_mix_df["Market Rent/Mo"] - unit_mix_df["Current Rent/Mo"]).clip(lower=0)).sum() * 12
            rent_upside_pct = ((loss_to_lease / gpr) * 100) if gpr > 0 else 0.0
            unit_mix_summary = unit_mix_df.copy()
            unit_mix_summary["Total SF"] = unit_mix_summary["Count"] * unit_mix_summary["Avg. SF"]
            unit_mix_summary["Monthly Rent"] = unit_mix_summary["Count"] * unit_mix_summary["Current Rent/Mo"]
            unit_mix_summary["Market Rent"] = unit_mix_summary["Count"] * unit_mix_summary["Market Rent/Mo"]
            unit_mix_summary = unit_mix_summary[["Unit Type", "Count", "Vacant Count", "Total SF", "Monthly Rent", "Market Rent"]]
            c1,c2 = st.columns(2)
            with c1: reimb = st.number_input("Expense Reimb. (NNN) ($)", min_value=0.0, value=0.0, step=5000.0, key="reimb_input_roll")
            with c2: vac_rate = st.number_input("Additional Economic/Credit Loss (%)", min_value=0.0, max_value=50.0, value=5.0, step=1.0, key="vr_roll")
        else:
            c1,c2,c3 = st.columns(3)
            with c1: gpr = st.number_input("Base Rent / GPR (annual $)", min_value=0.0, value=0.0, step=10000.0, key="gpr_input")
            with c2: reimb = st.number_input("Expense Reimb. (NNN) ($)", min_value=0.0, value=0.0, step=5000.0, key="reimb_input_quick")
            with c3: vac_rate = st.number_input("Vacancy & Credit Loss (%)", min_value=0.0, max_value=50.0, value=5.0, step=1.0, key="vr_quick")
        if use_unit_mix:
            credit_loss_base = max(gpr - physical_vacancy_loss + reimb, 0.0)
            credit_loss_amount = credit_loss_base * (vac_rate/100)
            egi = credit_loss_base - credit_loss_amount
        else:
            credit_loss_amount = (gpr + reimb) * (vac_rate/100)
            egi = (gpr + reimb) - credit_loss_amount
        market_total_annual = ((total_market_mo * 12) + reimb) if use_unit_mix else (gpr + reimb)
        st.markdown(f'<div class="mc"><div class="ml">Effective Gross Income</div><div class="mv" style="font-size:1.1rem;">{fmt_d(egi)}</div></div>', unsafe_allow_html=True)
        if use_unit_mix:
            st.markdown(f'<div class="mc blue"><div class="ml">Annual Loss to Lease</div><div class="mv blue" style="font-size:1.1rem;">{fmt_d(loss_to_lease)}</div><div class="rr"><span class="l">Rent Upside %</span><span class="v">{rent_upside_pct:.1f}%</span></div></div>', unsafe_allow_html=True)
        st.markdown(sl("Operating Expenses (Annual)"), unsafe_allow_html=True)
        use_opex_override = st.toggle("Use Expense Ratio Override (% of EGI)", key="opex_override_toggle", help="Enable when you do not have line-item expense data. Calculates total OpEx as a percentage of EGI. Typical ranges: Multifamily 35-45% | NNN Retail 10-20% | Gross Retail 30-40% | Office 35-50% | Industrial NNN 10-25% | Industrial Gross 30-40% | Mixed-Use 35-50% | Land 5-10%")
        if use_opex_override:
            opex_override_pct = st.number_input("Expense Ratio (% of EGI)", min_value=1.0, max_value=95.0, value=float(st.session_state.get("opex_override_pct", 35.0)), step=1.0, key="opex_override_pct")
            total_opex = egi * (opex_override_pct / 100.0) if egi > 0 else 0.0
            opex_ratio = opex_override_pct
            opex_tax = total_opex * 0.40
            opex_ins = total_opex * 0.30
            opex_mgmt_d = total_opex * 0.10
            opex_maint = total_opex * 0.10
            opex_util = total_opex * 0.05
            opex_res = total_opex * 0.05
            opex_oth = 0.0
            income_tax_line = opex_tax
            fixed_opex = total_opex
            fixed_opex_base = total_opex - opex_mgmt_d
            use_tax_reassessment = False
            st.markdown(f'<div class="co info"><strong>Expense Ratio Override Active:</strong> {opex_override_pct:.1f}% of EGI = <strong>{fmt_d(total_opex)}</strong>. Line items are distributed via placeholder percentages for PDF generation. <strong>Note:</strong> In Miami-Dade, taxes and insurance frequently represent a significantly higher burden (60-70%+) of gross expenses.</div>', unsafe_allow_html=True)
        else:
            c1,c2,c3 = st.columns(3)
            with c1: opex_tax = st.number_input("Property Taxes", min_value=0.0, value=0.0, step=1000.0, key="ot")
            use_tax_reassessment = st.toggle("Enable Institutional Tax Reassessment (Miami-Dade Standard)", key="tax_reassess")
            with c2: opex_ins = st.number_input("Insurance", min_value=0.0, value=0.0, step=1000.0, key="opex_ins_input")
            with c3: opex_mgmt_pct = st.number_input("Mgmt Fee (%)", min_value=0.0, max_value=15.0, value=5.0, step=0.5, key="om")
            tax_rate_pct = st.number_input("Pro-Forma Tax Rate (%)", min_value=0.0, max_value=10.0, value=float(st.session_state.get("tax_rate_pct", 2.10)), step=0.01, key="tax_rate_pct", help="Miami-Dade commercial millage is approximately 2.0% to 2.2% in 2026 depending on municipality. Verify current rate at miamidade.gov/pa for the subject's tax district.")
            c1,c2,c3 = st.columns(3)
            with c1: opex_maint = st.number_input("Maintenance", min_value=0.0, value=0.0, step=1000.0, key="omn")
            with c2: opex_util = st.number_input("Utilities (owner-paid)", min_value=0.0, value=0.0, step=500.0, key="ou")
            with c3: opex_res = st.number_input("Reserves", min_value=0.0, value=0.0, step=500.0, key="orr")
            opex_oth = st.number_input("Other Expenses", min_value=0.0, value=0.0, step=500.0, key="oox")
            mgmt_floor = 0.0 if subj_type == "Land" else ((subj_units * 250) if (subj_type == "Multifamily" and subj_units > 0) else (subj_sf * 0.10))
            opex_mgmt_d = max(egi*(opex_mgmt_pct/100), mgmt_floor)
            fixed_opex_base = opex_ins + opex_maint + opex_util + opex_res + opex_oth
            fixed_opex = fixed_opex_base + opex_mgmt_d
        loaded_cap = 0.0
        st.markdown(sl("Capitalization"), unsafe_allow_html=True)
        with st.expander("Miami-Dade Market Cap Rate Reference — Q1 2026"):
            st.markdown(
                """| Property Type | Cap Rate Range | Commentary |
| --- | --- | --- |
| Multifamily garden | 4.75% to 5.75% | compressed by rent growth |
| Multifamily mid-rise | 5.00% to 6.25% | |
| Retail Strip Center | 6.50% to 8.00% | vacancy-dependent |
| Single-Tenant NNN Retail | 5.50% to 7.00% | credit-tenant driven |
| Office Class A | 7.00% to 8.50% | elevated due to remote work |
| Office Class B and C | 8.50% to 11.00% | distressed pricing |
| Industrial and Warehouse | 5.00% to 6.50% | strong demand |
| Mixed-Use | 5.75% to 7.25% | ground-floor retail drives range |
| Land | Not applicable | sales comparison preferred |"""
            )
            st.caption("Source: CoStar, CBRE, and Marcus & Millichap Q1 2026 Miami-Dade market reports. For reference only.")
        c1,c2 = st.columns(2)
        with c1:
            inc_cap = st.number_input("Market Cap Rate (%)", min_value=0.1, max_value=20.0, value=6.5, step=0.25, key="ic")
            tenant_quality = st.selectbox("Anchor Tenant Credit Profile", ["Investment Grade (S&P/Moody's Rated)", "Regional / Strong Private", "Local / Unrated Mom & Pop"], key="tq_input")
            if tenant_quality == "Regional / Strong Private":
                risk_premium = 0.25
            elif tenant_quality == "Local / Unrated Mom & Pop":
                risk_premium = 0.75
            else:
                risk_premium = 0.0
            adjusted_cap = inc_cap + risk_premium
            if use_tax_reassessment:
                loaded_cap = adjusted_cap + tax_rate_pct
                noi_before_taxes = egi - fixed_opex
                income_value = (noi_before_taxes / (loaded_cap / 100)) if loaded_cap > 0 else 0
                income_tax_line = income_value * (tax_rate_pct / 100) if income_value else 0.0
                st.markdown(f'<div class="co info"><strong>Institutional tax mode active:</strong> loaded cap rate = <strong>{loaded_cap:.2f}%</strong> ({adjusted_cap:.2f}% market cap + {tax_rate_pct:.2f}% tax load), producing pro-forma taxes of <strong>{fmt_d(income_tax_line)}</strong>.</div>', unsafe_allow_html=True)
            else:
                income_tax_line = opex_tax
                income_value = ((egi - fixed_opex - income_tax_line) / (adjusted_cap / 100)) if adjusted_cap > 0 else 0
            total_opex = fixed_opex + income_tax_line
            noi = egi - total_opex
            opex_ratio = (total_opex/egi*100) if egi>0 else 0
            st.caption(f"Applied Risk Premium: +{risk_premium:.2f}% (Total Cap: {adjusted_cap:.2f}%) based on tenant credit-worthiness.")
        with c2:
            st.markdown(f'<div class="mc ok"><div class="ml">Value via Direct Cap</div><div class="mv green">{fmt_d(income_value)}</div></div>', unsafe_allow_html=True)
        m1,m2,m3 = st.columns(3)
        with m1: st.markdown(f'<div class="mc"><div class="ml">Total OpEx</div><div class="mv red" style="font-size:1.1rem;">{fmt_d(total_opex)}</div></div>', unsafe_allow_html=True)
        with m2: st.markdown(f'<div class="mc"><div class="ml">Expense Ratio</div><div class="mv" style="font-size:1.1rem;">{opex_ratio:.1f}%</div></div>', unsafe_allow_html=True)
        with m3:
            nc = "green" if noi>0 else "red"
            st.markdown(f'<div class="mc"><div class="ml">NOI</div><div class="mv {nc}" style="font-size:1.1rem;">{fmt_d(noi)}</div></div>', unsafe_allow_html=True)
        st.markdown(sl("Debt Service Coverage Ratio"), unsafe_allow_html=True)
        c1,c2,c3 = st.columns(3)
        with c1: loan_ltv = st.number_input("Assumed LTV (%)", min_value=0.0, max_value=100.0, value=65.0, step=1.0, key="dscr_ltv")
        with c2: loan_rate = st.number_input("Assumed Interest Rate (%)", min_value=0.0, max_value=20.0, value=7.25, step=0.25, key="dscr_rate")
        with c3: loan_amort_years = st.number_input("Amortization (Years)", min_value=1, max_value=40, value=25, step=1, key="dscr_amort")
        if income_value > 0 and loan_ltv == 0:
            st.markdown('<div class="co warn">LTV is set to 0% — DSCR and debt metrics are not calculated. Enter an assumed LTV to enable debt service analysis.</div>', unsafe_allow_html=True)
        if income_value > 0 and loan_ltv > 0 and loan_rate is not None and loan_amort_years:
            loan_amount = income_value * (loan_ltv / 100)
            monthly_rate = (loan_rate / 100) / 12
            total_payments = int(loan_amort_years * 12)
            if monthly_rate > 0 and total_payments > 0:
                monthly_payment = loan_amount * (monthly_rate * (1 + monthly_rate) ** total_payments) / (((1 + monthly_rate) ** total_payments) - 1)
            elif total_payments > 0:
                monthly_payment = loan_amount / total_payments
            else:
                monthly_payment = 0.0
            annual_debt_service = monthly_payment * 12
            dscr_value = (noi / annual_debt_service) if annual_debt_service > 0 else None
            if noi <= 0:
                dscr_value = None
                dscr_no_income = True
            else:
                dscr_no_income = False
            if dscr_no_income:
                st.markdown('<div class="mc alert"><div class="ml">Debt Service Coverage Ratio</div><div class="mv red" style="font-size:1.1rem;">N/A</div><div class="rr"><span class="l">Assessment</span><span class="v">NOI is zero or negative — property cannot support any debt at current income levels</span></div></div>', unsafe_allow_html=True)
            if dscr_value is not None:
                if dscr_value >= 1.25:
                    dscr_class, dscr_color, dscr_label = "ok", "green", "Strong coverage"
                elif dscr_value >= 1.20:
                    dscr_class, dscr_color, dscr_label = "gold", "gold", "Meets minimum coverage — monitor closely"
                else:
                    dscr_class, dscr_color, dscr_label = "alert", "red", "Below 1.20x — below lender minimums, limited buyer pool"
                st.markdown(
                    f'<div class="mc {dscr_class}"><div class="ml">Debt Service Coverage Ratio</div>'
                    f'<div class="mv {dscr_color}" style="font-size:1.2rem;">{dscr_value:.2f}x</div>'
                    f'<div class="rr"><span class="l">Annual Debt Service</span><span class="v">{fmt_d(annual_debt_service)}</span></div>'
                    f'<div class="rr"><span class="l">Loan Amount</span><span class="v">{fmt_d(loan_amount)}</span></div>'
                    f'<div class="rr"><span class="l">Assessment</span><span class="v">{dscr_label}</span></div></div>',
                    unsafe_allow_html=True,
                )
                if annual_debt_service > 0 and egi > 0:
                    fixed_opex_only = total_opex - opex_mgmt_d
                    mgmt_fee_rate = opex_mgmt_pct / 100.0
                    target_noi = annual_debt_service * 1.20
                    if (gpr + reimb) > 0:
                        if (1.0 - mgmt_fee_rate) > 0:
                            breakeven_egi_needed_pct = (target_noi + fixed_opex_only) / (1.0 - mgmt_fee_rate)
                            floor_fee_at_breakeven = mgmt_floor
                            pct_fee_at_breakeven = breakeven_egi_needed_pct * mgmt_fee_rate
                            if mgmt_floor > 0 and floor_fee_at_breakeven > pct_fee_at_breakeven:
                                breakeven_egi_needed = target_noi + fixed_opex_only + mgmt_floor
                            else:
                                breakeven_egi_needed = breakeven_egi_needed_pct
                        else:
                            breakeven_egi_needed = target_noi + fixed_opex_only + mgmt_floor
                        breakeven_occ_pct = (breakeven_egi_needed / (gpr + reimb)) * 100.0
                    else:
                        breakeven_occ_pct = None
                    if breakeven_occ_pct is not None:
                        if breakeven_occ_pct > 100.0:
                            st.markdown(
                                f'<div class="mc alert"><div class="ml">Break-Even Occupancy (at 1.20x DSCR)</div>'
                                f'<div class="mv red" style="font-size:1.1rem;">Unachievable</div>'
                                f'<div class="rr"><span class="l">Break-Even Required</span><span class="v">{breakeven_occ_pct:.1f}% of GPR</span></div>'
                                f'<div class="rr"><span class="l">Assessment</span><span class="v">NOI at full occupancy is insufficient to achieve 1.20x DSCR. This asset cannot support the modeled debt under any occupancy scenario.</span></div>'
                                f'</div>',
                                unsafe_allow_html=True
                            )
                            st.session_state.breakeven_occ = breakeven_occ_pct
                        elif subj_occ == 0.0:
                            st.markdown(
                                f'<div class="mc"><div class="ml">Break-Even Occupancy (at 1.20x DSCR)</div>'
                                f'<div class="mv" style="font-size:1.1rem;">{breakeven_occ_pct:.1f}%</div>'
                                f'<div class="rr"><span class="l">Assessment</span><span class="v">Enter subject occupancy in Tab 1 to enable comparison.</span></div>'
                                f'</div>',
                                unsafe_allow_html=True
                            )
                            st.session_state.breakeven_occ = breakeven_occ_pct
                        else:
                            beo_color = "green" if breakeven_occ_pct < subj_occ else "red"
                            beo_label = "Below current occupancy — adequate cushion" if breakeven_occ_pct < subj_occ else "Exceeds current occupancy — debt service at risk"
                            st.markdown(
                                f'<div class="mc"><div class="ml">Break-Even Occupancy (at 1.20x DSCR)</div>'
                                f'<div class="mv {beo_color}" style="font-size:1.1rem;">{breakeven_occ_pct:.1f}%</div>'
                                f'<div class="rr"><span class="l">Current Occupancy</span><span class="v">{subj_occ:.0f}%</span></div>'
                                f'<div class="rr"><span class="l">Assessment</span><span class="v">{beo_label}</span></div>'
                                f'</div>',
                                unsafe_allow_html=True
                            )
                            st.session_state.breakeven_occ = breakeven_occ_pct
        st.markdown(sl("Stabilized (optional)"), unsafe_allow_html=True)
        use_stab = st.toggle("Calculate stabilized value", key="us")
        if use_stab:
            c1,c2 = st.columns(2)
            with c1: stab_rent = st.number_input("Stabilized GPR+Reimb (annual)", min_value=0.0, value=float(market_total_annual if use_unit_mix else (gpr + reimb)), step=10000.0, key="sr")
            with c2: stab_occ = st.number_input("Stabilized Occ (%)", min_value=0.0, max_value=100.0, value=95.0, step=1.0, key="stoc")
            stab_egi = stab_rent*(stab_occ/100)
            stab_mgmt_fee = max(stab_egi * (opex_mgmt_pct/100), mgmt_floor)
            if use_tax_reassessment:
                stab_loaded_cap = adjusted_cap + tax_rate_pct
                stab_value = ((stab_egi - fixed_opex_base - stab_mgmt_fee) / (stab_loaded_cap/100)) if stab_loaded_cap>0 else 0
                stab_opex = fixed_opex_base + stab_mgmt_fee + (stab_value * (tax_rate_pct / 100))
            else:
                stab_opex = fixed_opex_base + stab_mgmt_fee + income_tax_line
                stab_value = ((stab_egi - fixed_opex_base - stab_mgmt_fee - income_tax_line) / (adjusted_cap/100)) if adjusted_cap>0 else 0
            stab_noi = stab_egi - stab_opex
            m1,m2 = st.columns(2)
            with m1: st.markdown(f'<div class="mc"><div class="ml">Stabilized NOI</div><div class="mv" style="font-size:1.1rem;">{fmt_d(stab_noi)}</div></div>', unsafe_allow_html=True)
            with m2: st.markdown(f'<div class="mc gold"><div class="ml">Stabilized Value</div><div class="mv gold">{fmt_d(stab_value)}</div></div>', unsafe_allow_html=True)
        st.markdown(sl("DCF (optional)"), unsafe_allow_html=True)
        use_dcf = st.toggle("Calculate DCF", key="ud")
        if use_dcf:
            c1,c2,c3,c4,c5 = st.columns(5)
            with c1: dy = st.selectbox("Hold Period", [3,5,7,10], index=1, key="dy")
            with c2: dg = st.number_input("Rev Growth (%)", min_value=0.0, max_value=10.0, value=3.0, step=0.5, key="dg")
            with c3: eg_rate = st.number_input("Exp Growth (%)", min_value=0.0, max_value=10.0, value=2.0, step=0.5, key="eg")
            with c4: dec = st.number_input("Exit Cap (%)", min_value=1.0, max_value=15.0, value=7.0, step=0.25, key="dec")
            with c5: ddr = st.number_input("Discount Rate (%)", min_value=1.0, max_value=20.0, value=8.5, step=0.5, key="ddr")
            use_mid_year_discounting = st.toggle("Use Mid-Year Discounting", key="mid_year_dcf")
            exit_selling_cost_pct = st.number_input("Exit Selling Cost (%)", min_value=0.0, max_value=10.0, value=4.0, step=0.5, key="exit_sc", help="Deducted from gross terminal value in the IRR calculation. Standard institutional range is 3% to 5% for brokerage, title, and closing costs.")
            base_egi = stab_egi if (use_stab and stab_egi > 0) else egi
            base_opex = stab_opex if (use_stab and stab_egi > 0) else total_opex
            cfs = []
            for yr in range(1,dy+1): cfs.append(base_egi*((1+dg/100)**yr) - base_opex*((1+eg_rate/100)**yr))
            tv_noi = base_egi*((1+dg/100)**(dy+1)) - base_opex*((1+eg_rate/100)**(dy+1))
            pv_operating_cf = 0.0
            for yr, cf in enumerate(cfs, 1):
                exponent = (yr - 0.5) if use_mid_year_discounting else yr
                pv_operating_cf += cf / ((1 + ddr/100) ** exponent)
            pv_terminal = ((tv_noi / (dec / 100)) / ((1 + ddr / 100) ** dy)) if dec > 0 else 0.0
            dcf_value = pv_operating_cf + pv_terminal
            st.markdown(f'<div class="mc blue"><div class="ml">DCF / NPV Value (Gross Reversion)</div><div class="mv blue">{fmt_d(dcf_value)}</div></div>', unsafe_allow_html=True)
            initial_investment = st.session_state.get("last_recon_val", 0.0)
            if initial_investment <= 0:
                initial_investment = income_value if income_value > 0 else 0.0
            irr_cash_flows = []
            if initial_investment > 0:
                irr_cash_flows = [-initial_investment]
                for cf in cfs[:-1]:
                    irr_cash_flows.append(cf)
                irr_cash_flows.append(cfs[-1] + ((tv_noi / (dec / 100)) * (1 - exit_selling_cost_pct / 100) if dec > 0 else 0.0))
                guess = 0.10
                converged = False
                for _ in range(1000):
                    npv = 0.0
                    d_npv = 0.0
                    for period, cf in enumerate(irr_cash_flows):
                        if use_mid_year_discounting and period > 0:
                            exponent = period - 0.5
                        else:
                            exponent = period
                        npv += cf / ((1 + guess) ** exponent)
                        if exponent > 0:
                            d_npv += (-exponent * cf) / ((1 + guess) ** (exponent + 1))
                    if abs(d_npv) < 1e-12:
                        break
                    next_guess = guess - (npv / d_npv)
                    if next_guess <= -0.999999:
                        break
                    if abs(next_guess - guess) < 0.00000001:
                        guess = next_guess
                        converged = True
                        break
                    guess = next_guess
                if converged and 0 <= guess <= 5:
                    irr_value = guess * 100
                    irr_class = "ok" if irr_value > ddr else "alert"
                    irr_color = "green" if irr_value > ddr else "red"
                    st.markdown(
                    f'<div class="mc {irr_class}"><div class="ml">Unlevered IRR vs Discount Rate (Net of Exit Costs)</div>'
                        f'<div class="mv {irr_color}" style="font-size:1.2rem;">{irr_value:.2f}%</div>'
                        f'<div class="rr"><span class="l">Discount Rate</span><span class="v">{ddr:.2f}%</span></div></div>',
                        unsafe_allow_html=True,
                    )
                elif converged and guess > 5:
                    st.markdown(
                        f'<div class="mc alert"><div class="ml">Unlevered IRR vs Discount Rate</div>'
                        f'<div class="mv red" style="font-size:1.2rem;">Result Exceeds 500%</div>'
                        f'<div class="rr"><span class="l">Assessment</span><span class="v">IRR converged above 500%. Verify acquisition basis, hold period, and exit cap inputs. This may indicate an underpriced entry or data entry error.</span></div></div>',
                        unsafe_allow_html=True,
                    )
                elif not converged:
                    st.markdown(
                        f'<div class="mc"><div class="ml">Unlevered IRR vs Discount Rate</div>'
                        f'<div class="mv" style="font-size:1.2rem;">No Solution</div>'
                        f'<div class="rr"><span class="l">Assessment</span><span class="v">IRR did not converge. The cash flows may not produce a solvable return. Check that the hold period, growth rates, and exit cap produce positive terminal value.</span></div></div>',
                        unsafe_allow_html=True,
                    )
    st.session_state.gpr = gpr
    st.session_state.reimb = reimb
    st.session_state.total_opex = total_opex
    st.session_state.inc_cap = inc_cap
    st.session_state.adjusted_cap = adjusted_cap
    st.session_state.risk_premium = risk_premium
    st.session_state.tenant_quality = tenant_quality
    st.session_state.opex_ins = opex_ins
    st.session_state.opex_mgmt_pct = opex_mgmt_pct
    st.session_state.fixed_opex_base = fixed_opex_base
    st.session_state.loss_to_lease = loss_to_lease
    st.session_state.physical_vacancy_loss = physical_vacancy_loss
    st.session_state.vac_rate = vac_rate
    st.session_state.mgmt_floor = mgmt_floor
    st.session_state.use_tax_reassessment = use_tax_reassessment
    st.session_state.income_tax_line = income_tax_line
    st.session_state.subj_occ = subj_occ
    st.session_state.dscr = dscr_value
    st.session_state.annual_debt_service = annual_debt_service
    st.session_state.loan_amount = loan_amount
    st.session_state.loan_ltv = loan_ltv
    st.session_state.loan_rate = loan_rate
    st.session_state.loan_amort_years = loan_amort_years
    st.session_state.irr_value = irr_value
    st.session_state.ddr = ddr if use_dcf else st.session_state.get("ddr", 8.5)
    st.session_state.use_dcf = use_dcf
    st.session_state.dscr_no_income = dscr_no_income
    st.session_state.breakeven_occ = breakeven_occ_pct

with tab4:
    st.markdown('<div class="co info">Optional \u2014 most relevant for newer or special-use properties.</div>', unsafe_allow_html=True)
    cost_value = 0; cost_land = 0; tot_repl = 0; td = 0; dep_imp = 0
    use_cost = st.toggle("Include Cost Approach", key="uc")
    if use_cost:
        st.markdown(sl("Land Value"), unsafe_allow_html=True)
        cost_land = st.number_input("Land Value Estimate ($)", min_value=0.0, value=0.0, step=25000.0, key="cla")
        if subj_type == "Land":
            tot_repl = 0.0
            dep_imp = 0.0
            td = 0.0
            cost_value = cost_land
        else:
            st.markdown(sl("Replacement Cost"), unsafe_allow_html=True)
            c1,c2 = st.columns(2)
            cost_by_unit = subj_type == "Multifamily" and subj_sf == 0 and subj_units > 0
            with c1: cost_psf = st.number_input("Replacement Cost per Unit ($)" if cost_by_unit else "Replacement Cost/SF ($)", min_value=0.0, value=0.0, step=10.0 if not cost_by_unit else 1000.0, key="cpsf")
            with c2:
                tot_repl = cost_psf*(subj_units if cost_by_unit else subj_sf)
                st.markdown(f'<div class="mc"><div class="ml">Total Replacement</div><div class="mv" style="font-size:1.1rem;">{fmt_d(tot_repl)}</div></div>', unsafe_allow_html=True)
            st.markdown(sl("Depreciation"), unsafe_allow_html=True)
            c1,c2,c3 = st.columns(3)
            with c1: dp = st.number_input("Physical (%)", min_value=0.0, max_value=100.0, value=0.0, step=5.0, key="dp")
            with c2: df = st.number_input("Functional (%)", min_value=0.0, max_value=100.0, value=0.0, step=5.0, key="df")
            with c3: de = st.number_input("External (%)", min_value=0.0, max_value=100.0, value=0.0, step=5.0, key="de")
            td_raw = dp + df + de
            if td_raw > 100.0:
                st.markdown('<div class="co warn"><strong>Depreciation Cap Applied:</strong> Physical, Functional, and External depreciation inputs sum to more than 100%. Total depreciation has been capped at 100%. Improvements cannot depreciate below zero value. Review your inputs.</div>', unsafe_allow_html=True)
            td = min(td_raw, 100.0)
            dep_imp = tot_repl * (1 - td / 100)
            cost_value = cost_land + dep_imp
        if subj_type == "Land":
            st.markdown(f'<div class="mc ok"><div class="ml">Cost Approach Value</div><div class="mv green">{fmt_d(cost_value)}</div></div>', unsafe_allow_html=True)
        else:
            m1,m2,m3 = st.columns(3)
            with m1: st.markdown(f'<div class="mc"><div class="ml">Total Depreciation</div><div class="mv" style="font-size:1.1rem;">{td:.0f}%</div></div>', unsafe_allow_html=True)
            with m2: st.markdown(f'<div class="mc"><div class="ml">Depreciated Improvements</div><div class="mv" style="font-size:1.1rem;">{fmt_d(dep_imp)}</div></div>', unsafe_allow_html=True)
            with m3: st.markdown(f'<div class="mc ok"><div class="ml">Cost Approach Value</div><div class="mv green">{fmt_d(cost_value)}</div></div>', unsafe_allow_html=True)

with tab5:
    st.markdown(sl("Value Indications"), unsafe_allow_html=True)
    appr = {}
    if sales_comp_value > 0: appr["Sales Comparison"] = sales_comp_value
    if income_value > 0: appr["Income Capitalization"] = income_value
    if cost_value > 0: appr["Cost Approach"] = cost_value
    if appr:
        for nm,vl in appr.items(): st.markdown(f'<div class="mc"><div class="ml">{nm}</div><div class="mv" style="font-size:1.2rem;">{fmt_d(vl)}</div></div>', unsafe_allow_html=True)
    else: st.markdown('<div class="co warn">No approaches have data yet.</div>', unsafe_allow_html=True)
    if st.session_state.get("weight_error", False):
        st.error(weight_error_message or f"CRITICAL: Manual weights sum to {total_manual_weight:.1f}%. They must sum to exactly 100% to generate a report.")
    st.markdown(sl("Sensitivity Analysis"), unsafe_allow_html=True)
    sensitivity_matrix = np.array([])
    sensitivity_cap_steps = np.array([])
    sensitivity_occ_steps = np.array([])
    sync_gpr = st.session_state.get("gpr", 0.0)
    sync_reimb = st.session_state.get("reimb", 0.0)
    sync_inc_cap = st.session_state.get("inc_cap", 0.0)
    sync_adjusted_cap = st.session_state.get("adjusted_cap", sync_inc_cap)
    sync_subj_occ = st.session_state.get("subj_occ", 0.0)
    sync_mgmt_pct = st.session_state.get("opex_mgmt_pct", 0.0)
    sync_fixed_opex_base = st.session_state.get("fixed_opex_base", 0.0)
    sync_loss_to_lease = st.session_state.get("loss_to_lease", 0.0)
    sync_physical_vacancy_loss = st.session_state.get("physical_vacancy_loss", 0.0)
    sync_vac_rate = st.session_state.get("vac_rate", 5.0)
    sync_mgmt_floor = st.session_state.get("mgmt_floor", 0.0)
    sync_tax_rate = st.session_state.get("tax_rate_pct", 2.10)
    sync_use_tax_reassessment = st.session_state.get("use_tax_reassessment", False)
    sync_income_tax_line = st.session_state.get("income_tax_line", 0.0)
    if sync_adjusted_cap > 0 and (sync_gpr + sync_reimb) > 0:
        sensitivity_cap_steps = np.array([sync_adjusted_cap - 0.5, sync_adjusted_cap, sync_adjusted_cap + 0.5], dtype=float)
        base_occ = max(min(100.0 - sync_vac_rate, 95.0), 5.0)
        raw_steps = np.array([base_occ - 5.0, base_occ, base_occ + 5.0], dtype=float)
        sensitivity_occ_steps = np.clip(raw_steps, 0.0, 100.0)
        sensitivity_occ_steps = np.unique(sensitivity_occ_steps)
        if len(sensitivity_occ_steps) < 3:
            sensitivity_occ_steps = np.array([max(base_occ - 5.0, 0.0), base_occ, min(base_occ + 5.0, 100.0)], dtype=float)
        sensitivity_rows = []
        for cap_step in sensitivity_cap_steps:
            row = []
            for occ_step in sensitivity_occ_steps:
                gross_collectible = sync_gpr - sync_physical_vacancy_loss - sync_loss_to_lease + sync_reimb
                sim_credit_loss = gross_collectible * (1 - (occ_step / 100))
                sim_egi = gross_collectible - sim_credit_loss
                sim_mgmt_fee = max(sim_egi * (sync_mgmt_pct / 100), sync_mgmt_floor)
                sim_noi = sim_egi - sync_fixed_opex_base - sim_mgmt_fee
                if sync_use_tax_reassessment:
                    cap_denom = (cap_step + sync_tax_rate) / 100
                else:
                    sim_noi -= sync_income_tax_line
                    cap_denom = cap_step / 100
                row.append((sim_noi / cap_denom) if cap_denom > 0 else 0.0)
            sensitivity_rows.append(row)
        sensitivity_matrix = np.array(sensitivity_rows, dtype=float)
        sensitivity_df = pd.DataFrame(
            sensitivity_matrix,
            index=[f"{cap_step:.2f}% Cap" for cap_step in sensitivity_cap_steps],
            columns=[f"{occ_step:.0f}% Collection" for occ_step in sensitivity_occ_steps],
        )
        st.dataframe(sensitivity_df.style.format("${:,.0f}"), use_container_width=True)
        st.caption("Collection rate represents the percentage of gross collectible income realized after credit loss and vacancy variation. A 5% shift is applied in each direction from the current effective rate.")
    else:
        st.markdown('<div class="co">Sensitivity analysis will populate after gross income and cap rate inputs are entered.</div>', unsafe_allow_html=True)
    st.markdown(sl("Weight Assignment"), unsafe_allow_html=True)
    st.markdown('<div class="co info"><strong>Tip:</strong> Income highest for investment properties, Sales highest for owner-user.</div>', unsafe_allow_html=True)
    wts = {}
    if appr:
        cols = st.columns(len(appr))
        for idx,(nm,vl) in enumerate(appr.items()):
            with cols[idx]:
                dw = round(100/len(appr))
                if idx == len(appr)-1: dw = 100 - dw*(len(appr)-1)
                wts[nm] = st.number_input(f"{nm} (%)", min_value=0.0, max_value=100.0, value=float(dw), step=5.0, key=f"wt_{nm}")
        twt = sum(wts.values())
        if abs(twt-100) > 0.5: st.markdown(f'<div class="co warn">Weights = <strong>{twt:.0f}%</strong> (must = 100%)</div>', unsafe_allow_html=True)
        valuation_weight_error = abs(twt-100) > 0.0001
        manual_weight_error = manual_weight_override and abs(total_manual_weight-100.0) > 0.01
        st.session_state.weight_error = manual_weight_error or valuation_weight_error
    else:
        twt = 0.0
        st.session_state.weight_error = manual_weight_override and abs(total_manual_weight-100.0) > 0.01
    pdf_button_disabled = st.session_state.get("weight_error", False) or abs(twt-100) > 0.5
    liquidation_discount = st.number_input("Liquidation Discount (%)", min_value=0.0, max_value=100.0, value=20.0, step=1.0, help="Liquidation discount applied to the As-Is value. A 20% discount reflects a typical forced liquidation scenario. Adjust based on asset type, market depth, and bank disposition mandate.", key="liq_disc")
    recon_val = sum((wts.get(n,0)/100)*v for n,v in appr.items()) if appr else 0
    st.markdown(sl("Recommendation"), unsafe_allow_html=True)
    capex_deduct = st.number_input("Less: Deferred Maintenance / CapEx ($)", min_value=0.0, value=0.0, step=5000.0, key="capex_deduct")
    as_is_value = recon_val - capex_deduct
    st.session_state.last_recon_val = as_is_value
    if as_is_value > 0: st.markdown(f'<div class="mc gold"><div class="ml">As-Is Reconciled Value</div><div class="mv gold" style="font-size:1.8rem;">{fmt_d(as_is_value)}</div></div>', unsafe_allow_html=True)
    c1,c2 = st.columns(2)
    with c1: rec_price = st.number_input("Recommended List Price ($)", min_value=0.0, value=float(max(as_is_value,0)), step=25000.0, key="rp")
    liquidation_value = max(as_is_value, 0) * (1 - (liquidation_discount / 100))
    with c1:
        st.markdown(f'<div class="mc alert"><div class="ml">Liquidation Value ({liquidation_discount:.0f}% Discount)</div><div class="mv red" style="font-size:1.5rem;">{fmt_d(liquidation_value)}</div></div>', unsafe_allow_html=True)
        st.caption(f"Liquidation value indicates a protective floor price of {fmt_d(liquidation_value)} after a {liquidation_discount:.0f}% discount from the As-Is reconciled value.")
    with c2: rec_just = st.text_input("Justification (if different)", key="rj")
    if rec_price > 0 and st.session_state.get("conf_score", "N/A") != "N/A":
        _cov_now = st.session_state.get("cov", 0.0)
        _band_pct = 0.05 if _cov_now < 10.0 else (0.10 if _cov_now <= 15.0 else 0.15)
        st.markdown(f'<div class="mc"><div class="rr"><span class="l">Price Range (\u00b1{_band_pct:.0%} based on {st.session_state.get("conf_score","N/A")} confidence)</span><span class="v gold">{fmt_d(rec_price*(1-_band_pct))} \u2014 {fmt_d(rec_price*(1+_band_pct))}</span></div></div>', unsafe_allow_html=True)
    elif rec_price > 0 and st.session_state.get("conf_score", "N/A") == "N/A":
        st.markdown('<div class="co warn">Price range requires comparable sales data. Enter and analyze comps in Tab 2 to generate a statistically supported price range.</div>', unsafe_allow_html=True)
    c1,c2 = st.columns(2)
    with c1: rec_mkt = st.selectbox("Marketing Time", MKT_TIMES, key="rm")
    with c2: rec_dom = st.number_input("Est. Days to Close (Distressed Disposition)", min_value=0, value=60, step=15, key="rdom")
    rec_narr = st.text_area("Broker Recommendation", key="rn", height=100, placeholder="2-3 sentences on disposition strategy...")
    recon_narrative = st.text_area("Reconciliation Narrative", key="recon_narrative_key", height=100, placeholder="Explain the weighting logic between the Sales and Income approaches...")
    st.markdown(sl("Market Overview (for PDF)"), unsafe_allow_html=True)
    c1,c2 = st.columns(2)
    with c1: mkt_vac = st.number_input("Submarket Vacancy (%)", min_value=0.0, max_value=50.0, value=0.0, step=1.0, key="mv2")
    with c2: mkt_abs = st.text_input("Absorption Trends", key="ma")
    mkt_factors = st.text_area("Market Factors", key="mf", height=80)
    st.markdown(sl("Demographics & Aerial Map"), unsafe_allow_html=True)
    st.markdown('<div class="co info"><strong>Free Data Sources:</strong> Use these links to quickly pull data for the fields below.</div>', unsafe_allow_html=True)
    link_c1, link_c2 = st.columns(2)
    with link_c1:
        st.link_button("↗ Free Demographics (Census Reporter)", url="https://censusreporter.org/", use_container_width=True)
    with link_c2:
        st.link_button("↗ Free Traffic Counts (Florida DOT)", url="https://tdaappsprod.dot.state.fl.us/fto/", use_container_width=True)
    c1,c2,c3 = st.columns(3)
    with c1: mkt_pop_1m = st.number_input("Population (1-Mile)", min_value=0, value=0, step=1000, key="mkt_pop_1m")
    with c2: mkt_pop_3m = st.number_input("Population (3-Mile)", min_value=0, value=0, step=1000, key="mkt_pop_3m")
    with c3: mkt_pop_5m = st.number_input("Population (5-Mile)", min_value=0, value=0, step=1000, key="mkt_pop_5m")
    c1,c2,c3 = st.columns(3)
    with c1: mkt_med_inc_1m = st.number_input("Median Income (1-Mi)", min_value=0, value=0, step=5000, key="mkt_inc_1m")
    with c2: mkt_med_inc_3m = st.number_input("Median Income (3-Mi)", min_value=0, value=0, step=5000, key="mkt_inc_3m")
    with c3: mkt_med_inc_5m = st.number_input("Median Income (5-Mi)", min_value=0, value=0, step=5000, key="mkt_inc_5m")
    c1,c2,c3 = st.columns(3)
    with c1: mkt_vpd_1m = st.number_input("Traffic VPD (1-Mi)", min_value=0, value=0, step=1000, key="mkt_vpd_1m")
    with c2: mkt_vpd_3m = st.number_input("Traffic VPD (3-Mi)", min_value=0, value=0, step=1000, key="mkt_vpd_3m")
    with c3: mkt_vpd_5m = st.number_input("Traffic VPD (5-Mi)", min_value=0, value=0, step=1000, key="mkt_vpd_5m")
    mkt_map_photo = st.file_uploader("Upload Aerial Map (CoStar/Google Earth)", type=['jpg', 'png', 'jpeg'], key="mkt_map")
    st.markdown(sl("Qualitative SWOT Analysis"), unsafe_allow_html=True)
    c1,c2 = st.columns(2)
    with c1:
        swot_strengths = st.text_area("Strengths", key="swot_strengths", height=110, placeholder="S: High visibility corner location")
        swot_opportunities = st.text_area("Opportunities", key="swot_opportunities", height=110, placeholder="O: Below-market rents create mark-to-market upside")
    with c2:
        swot_weaknesses = st.text_area("Weaknesses", key="swot_weaknesses", height=110, placeholder="W: Deferred maintenance on mechanicals")
        swot_threats = st.text_area("Threats", key="swot_threats", height=110, placeholder="T: New competing supply and tenant rollover risk")

    # â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
    # PDF GENERATION
    # â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
    st.markdown('<div class="dv"></div>', unsafe_allow_html=True)

    def gen_pdf():
        buf = io.BytesIO()
        inst_mode = st.session_state.get("inst_mode", False)
        NV=HexColor("#002147"); SV=HexColor("#94A3B8"); LG=HexColor("#F8FAFC")
        BG2=HexColor("#E2E8F0"); MG=HexColor("#64748B"); TXT=HexColor("#334155")
        LIGHT_GRID = HexColor("#E2E8F0")
        report_title_firm = "ESTUPINAN GROUP | COMMERCIAL ADVISORY" if inst_mode else BRAND_NAME.upper()
        display_brand_name = "Estupinan Group | Commercial Advisory" if inst_mode else BRAND_NAME
        signature_brand_name = "Estupinan Group | Commercial Advisory" if inst_mode else BRAND_NAME
        signature_broker_firm = "Estupinan Group | Commercial Advisory" if inst_mode else BROKER_FIRM
        analyst_name = st.session_state.get("analyst_name", "Madelin Estupinan")
        table_grid_width = 0.1
        table_grid_color = HexColor("#E2E8F0")
        comp_header_padding = 3 if inst_mode else 5
        comp_body_padding = 3 if inst_mode else 4
        summary_padding = 3 if inst_mode else 6
        sty = getSampleStyleSheet()
        sb = ParagraphStyle("b",parent=sty["Normal"],fontSize=9,leading=9.9,textColor=TXT,fontName="Helvetica")
        sbs = ParagraphStyle("bs",parent=sb,fontSize=7.5,leading=8.25)
        ssec = ParagraphStyle("sec",parent=sb,fontSize=10,leading=11,textColor=NV,spaceBefore=8,spaceAfter=3,fontName="Helvetica-Bold")
        ssm = ParagraphStyle("sm",parent=sty["Normal"],fontSize=7,leading=7.7,textColor=MG)
        W = 7.4*inch  # usable width

        def Spacer(width, height):
            return RLSpacer(width, height * 0.6)

        def draw_page_chrome(canvas, doc):
            if inst_mode:
                canvas.saveState()
                canvas.setFillColor(HexColor("#CBD5E1"))
                canvas.setFont("Helvetica-Bold", 36)
                canvas.translate(letter[0] / 2, letter[1] / 2)
                canvas.rotate(45)
                canvas.drawCentredString(0, 0, "CONFIDENTIAL")
                canvas.restoreState()
                canvas.saveState()
                canvas.setFont("Helvetica", 6)
                canvas.setFillColor(MG)
                canvas.drawCentredString(letter[0] / 2, 0.22 * inch, "This BOV is prepared exclusively for the stated intended user and is not a certified appraisal. Not for third-party reliance. © Estupinan Group | First Service Realty by ERA.")
                canvas.restoreState()
            canvas.saveState()
            canvas.setFont("Helvetica", 6)
            canvas.setFillColor(MG)
            canvas.drawRightString(letter[0] - 0.3*inch, 0.22 * inch, f"Page {canvas.getPageNumber()}")
            canvas.restoreState()

        def band(txt):
            t=Table([[Paragraph(f"<b>{txt}</b>",ParagraphStyle("hb",parent=sb,fontSize=10,textColor=white,fontName="Helvetica-Bold"))]],colWidths=[W])
            t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),NV),("TOPPADDING",(0,0),(-1,-1),6),("BOTTOMPADDING",(0,0),(-1,-1),6),("LEFTPADDING",(0,0),(-1,-1),10)]))
            return t

        def build_pdf_image(uploaded_file, target_width, max_height=None):
            if uploaded_file is None:
                return None
            try:
                img_bytes = uploaded_file.getvalue()
                with PILImage.open(io.BytesIO(img_bytes)) as pil_img:
                    pil_img = pil_img.convert("RGB")
                    img_width, img_height = pil_img.size
                    render_stream = io.BytesIO()
                    pil_img.save(render_stream, format="PNG")
                    render_bytes = render_stream.getvalue()
                if img_width <= 0 or img_height <= 0:
                    return None
                aspect_ratio = img_height / img_width
                render_width = target_width
                render_height = render_width * aspect_ratio
                if max_height and render_height > max_height:
                    render_height = max_height
                    render_width = render_height / aspect_ratio
                return Image(io.BytesIO(render_bytes), width=render_width, height=render_height)
            except Exception:
                return None

        def expert_signature_table():
            left_block = Paragraph(
                f"<b>Madelin Estupinan</b><br/>Principal/Founder of Estupinan Group<br/>First Service Realty by ERA<br/>CCIM 101 Graduate 2005 | CB Commercial 2005<br/>Lic. {LICENSE_AGENT} | Phone: (786) 514-8046 | Email: madelin@estupinangroup.com",
                sbs
            )
            eric_block = Paragraph(
                "<b>Eric Yi</b><br/>Business Analyst of Estupinan Group<br/>First Service Realty by ERA<br/>Commercial Associate | CCIM 101 Graduate 2024<br/>Phone: (786) 897-5684 | Email: eric@estupinangroup.com",
                sbs
            )
            eddie_block = Paragraph(
                "<b>Eddie San Roman</b><br/>Broker of Record | 24 Asset Management Corp<br/>First Service Realty by ERA<br/>Licensed Broker | BK3672701<br/>Email: eddie@estupinangroup.com",
                sbs
            )
            if analyst_name == "Eddie San Roman":
                cols = [left_block, eddie_block]
            elif analyst_name == "Eric Yi":
                cols = [left_block, eric_block]
            else:
                cols = [left_block, eric_block]
            sig = Table([cols], colWidths=[3.6*inch, 3.6*inch])
            sig.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"TOP"),("TOPPADDING",(0,0),(-1,-1),25)]))
            return sig

        story = []
        if use_letter:
            recipient_display_name = recipient_name or "Client Representative"
            recipient_display_title = recipient_title or prep_for or "Client Representative"
            transmittal_table = Table([
                ["Opinion of Market Value", fmt_d(rec_price)],
                [f"Liquidation Value ({liquidation_discount:.0f}% Discount)", fmt_d(liquidation_value)],
            ], colWidths=[3.5*inch, 2.2*inch])
            transmittal_table.setStyle(TableStyle([
                ("GRID",(0,0),(-1,-1),table_grid_width,table_grid_color),
                ("FONTNAME",(0,0),(0,-1),"Helvetica-Bold"),
                ("FONTNAME",(1,0),(1,-1),"Helvetica-Bold"),
                ("TEXTCOLOR",(0,0),(0,-1),MG),
                ("TEXTCOLOR",(1,0),(1,-1),HexColor("#0B1D3A")),
                ("ALIGN",(1,0),(1,-1),"RIGHT"),
                ("TOPPADDING",(0,0),(-1,-1),6),
                ("BOTTOMPADDING",(0,0),(-1,-1),6),
                ("LEFTPADDING",(0,0),(-1,-1),12),
                ("RIGHTPADDING",(0,0),(-1,-1),12),
            ]))
            top_bar = Table([[""]], colWidths=[W], rowHeights=[4])
            top_bar.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),SV)]))
            story.append(top_bar)
            story.append(Spacer(1, 0.55*inch))
            story.append(Paragraph(report_title_firm, ParagraphStyle("ltr_firm", parent=sty["Normal"], fontSize=9 if inst_mode else 8, textColor=MG, alignment=TA_CENTER, spaceAfter=6, fontName="Helvetica-Bold" if inst_mode else "Helvetica")))
            story.append(Paragraph("TRANSMITTAL LETTER", ParagraphStyle("ltr_title", parent=sty["Normal"], fontSize=20, textColor=NV, fontName="Helvetica-Bold", alignment=TA_CENTER, spaceAfter=20)))
            story.append(HRFlowable(width="36%", thickness=1.5, color=SV, spaceAfter=22))
            story.append(Paragraph(rpt_date.strftime("%B %d, %Y") if rpt_date else datetime.now().strftime("%B %d, %Y"), ParagraphStyle("ltr_date", parent=sb, alignment=TA_RIGHT, textColor=MG, spaceAfter=18)))
            story.append(Paragraph(sanitize_pdf_text(recipient_display_name), ParagraphStyle("ltr_rec_name", parent=sb, fontName="Helvetica-Bold", spaceAfter=2)))
            story.append(Paragraph(sanitize_pdf_text(recipient_display_title), ParagraphStyle("ltr_rec_title", parent=sb, textColor=MG, spaceAfter=2)))
            if prep_for and prep_for != recipient_display_title:
                story.append(Paragraph(sanitize_pdf_text(prep_for), ParagraphStyle("ltr_rec_bank", parent=sb, textColor=MG, spaceAfter=2)))
            story.append(Spacer(1, 0.18*inch))
            story.append(Paragraph(f"Dear {sanitize_pdf_text(recipient_display_name)},", ParagraphStyle("ltr_salute", parent=sb, spaceAfter=12)))
            story.append(Paragraph(sanitize_pdf_text(engagement_narrative), ParagraphStyle("ltr_body", parent=sb, alignment=TA_JUSTIFY, leading=14, spaceAfter=14)))
            story.append(Paragraph(
                "Based on our review of current Miami-Dade market evidence, comparable sale activity, and income profile, we respectfully submit the following value conclusions for your consideration:",
                ParagraphStyle("ltr_intro", parent=sb, alignment=TA_JUSTIFY, leading=14, spaceAfter=10)
            ))
            story.append(transmittal_table)
            story.append(Spacer(1, 0.3*inch))
            story.append(Paragraph("Sincerely,", ParagraphStyle("ltr_close", parent=sb, spaceAfter=14)))
            story.append(expert_signature_table())
            story.append(PageBreak())

        # â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
        # PAGE 1: COVER
        # â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
        # Top gold accent bar
        top_bar = Table([[""]],colWidths=[W],rowHeights=[4])
        top_bar.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),SV)]))
        story.append(top_bar)

        story.append(Spacer(1, 0.5*inch))

        # Firm name
        story.append(Paragraph(report_title_firm, ParagraphStyle("firm", parent=sty["Normal"], fontSize=9 if st.session_state.inst_mode else 8, textColor=MG, alignment=TA_CENTER, spaceAfter=6, fontName="Helvetica-Bold" if st.session_state.inst_mode else "Helvetica")))

        # Title
        story.append(Paragraph("Broker Opinion of Value", ParagraphStyle("title",parent=sty["Normal"],fontSize=28,textColor=NV,fontName="Helvetica-Bold",alignment=TA_CENTER,spaceAfter=28)))

        # Gold divider - wider
        story.append(HRFlowable(width="40%",thickness=2,color=SV,spaceAfter=28))

        # Property address
        story.append(Paragraph(sanitize_pdf_text((subj_address or "Subject Property").upper()), ParagraphStyle("addr",parent=sty["Normal"],fontSize=15,textColor=black,alignment=TA_CENTER,fontName="Helvetica-Bold",spaceAfter=16)))

        # Subtitle
        parts = [subj_type]
        if subj_sf: parts.append(f"{subj_sf:,} SF")
        if subj_yr: parts.append(f"Built {subj_yr}")
        if subj_zone: parts.append(sanitize_pdf_text(subj_zone))
        story.append(Paragraph("  |  ".join(parts), ParagraphStyle("sub",parent=sty["Normal"],fontSize=10,textColor=MG,alignment=TA_CENTER,spaceAfter=40)))
        cover_subject_photo = build_pdf_image(subj_photo, 5.5*inch, max_height=3.0*inch)
        if cover_subject_photo is not None:
            subject_photo_wrap = Table([[cover_subject_photo]], colWidths=[W])
            subject_photo_wrap.setStyle(TableStyle([
                ("ALIGN",(0,0),(-1,-1),"CENTER"),
                ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
                ("LEFTPADDING",(0,0),(-1,-1),6),
                ("RIGHTPADDING",(0,0),(-1,-1),6),
                ("TOPPADDING",(0,0),(-1,-1),6),
                ("BOTTOMPADDING",(0,0),(-1,-1),6),
            ]))
            story.append(subject_photo_wrap)
            story.append(Spacer(1, 0.25*inch))
        executive_units_sf = subj_units if subj_units > 0 else subj_sf
        executive_units_label = "Total Units" if subj_units > 0 else "Rentable Area (SF)"
        executive_summary = Table([
            ["Asset Class", subj_type or "N/A", executive_units_label, f"{executive_units_sf:,}" if executive_units_sf else "N/A"],
            ["Year Built", str(subj_yr) if subj_yr else "N/A", "Occupancy", f"{subj_occ:.0f}%"],
            ["Last Sale Price", fmt_d(last_sale_price) if (last_sale_price is not None and last_sale_price != 0) else ("$0 — Deed-in-Lieu / No Consideration" if (last_sale_price == 0 and last_sale_date) else "N/A"), "Last Sale Date", last_sale_date.strftime("%m/%Y") if last_sale_date else "N/A"],
        ], colWidths=[1.2*inch, 2.35*inch, 1.2*inch, 2.65*inch])
        executive_summary.setStyle(TableStyle([
            ("GRID",(0,0),(-1,-1),table_grid_width,table_grid_color),
            ("ROWBACKGROUNDS",(0,0),(-1,-1),[white, LG]),
            ("TOPPADDING",(0,0),(-1,-1),4),
            ("BOTTOMPADDING",(0,0),(-1,-1),4),
            ("LEFTPADDING",(0,0),(-1,-1),6),
            ("RIGHTPADDING",(0,0),(-1,-1),6),
            ("FONTNAME",(0,0),(0,-1),"Helvetica-Bold"),
            ("FONTNAME",(2,0),(2,-1),"Helvetica-Bold"),
            ("TEXTCOLOR",(0,0),(0,-1),MG),
            ("TEXTCOLOR",(2,0),(2,-1),MG),
            ("ALIGN",(1,0),(1,-1),"RIGHT"),
            ("ALIGN",(3,0),(3,-1),"RIGHT"),
        ]))
        story.append(KeepTogether([Paragraph("<b>EXECUTIVE SUMMARY</b>", ssec), Spacer(1,4), executive_summary]))
        story.append(Spacer(1,0.3*inch))

        # Value box â€” single cell, vertically centered
        value_box_style = ParagraphStyle("vbox", parent=sb, alignment=TA_CENTER, leading=10, textColor=white, fontName="Helvetica-Bold")
        price_box_style = ParagraphStyle("pbox", parent=value_box_style, leading=24)
        olv_price_style = ParagraphStyle("obox", parent=value_box_style, leading=14)
        vb_inner = Table([
            [Paragraph("<font size='9' color='#8B9AAF'>OPINION OF MARKET VALUE</font>", value_box_style)],
            [Paragraph(f"<font size='22'>{fmt_d(rec_price)}</font>", price_box_style)],
            [Spacer(1, 4)],
            [Paragraph(f"<font size='9' color='#E5534B'>LIQUIDATION VALUE ({liquidation_discount:.0f}% DISCOUNT)</font>", value_box_style)],
            [Paragraph(f"<font size='14' color='#E5534B'>{fmt_d(liquidation_value)}</font>", olv_price_style)],
        ], colWidths=[5*inch])
        vb_inner.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,-1),NV),
            ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
            ("ALIGN",(0,0),(-1,-1),"CENTER"),
            ("BOX",(0,0),(-1,-1),1.5,SV),
            ("TOPPADDING",(0,0),(-1,-1),3),
            ("BOTTOMPADDING",(0,0),(-1,-2),3),
            ("BOTTOMPADDING",(0,-1),(-1,-1),14),
        ]))
        # Center the box
        vb_wrap = Table([[vb_inner]], colWidths=[W])
        vb_wrap.setStyle(TableStyle([("ALIGN",(0,0),(-1,-1),"CENTER"),("LEFTPADDING",(0,0),(-1,-1),0),("RIGHTPADDING",(0,0),(-1,-1),0)]))
        story.append(vb_wrap)

        story.append(Spacer(1, 0.05*inch))

        # Prepared details
        pdf_effective_date = st.session_state.get("effective_date")
        pdf_intended_use = st.session_state.get("intended_use") or "Not specified"
        pdf_intended_user = st.session_state.get("intended_user") or "Not specified"
        cov_data = [
            ["Prepared For:", prep_for or "Client Representative"],
            ["Prepared By:", analyst_name],
            ["Firm:", display_brand_name],
            ["Contact:", BRAND_CONTACT],
            ["Date:", rpt_date.strftime("%B %d, %Y") if rpt_date else "\u2014"],
            ["Effective Date:", pdf_effective_date.strftime("%B %d, %Y") if isinstance(pdf_effective_date, date) else "Not specified"],
            ["Intended Use / User:", f"{pdf_intended_use} | {pdf_intended_user}"],
        ]
        ct2=Table(cov_data,colWidths=[1.1*inch,5.9*inch])
        ct2.setStyle(TableStyle([
            ("FONTNAME",(0,0),(0,-1),"Helvetica-Bold"),
            ("FONTSIZE",(0,0),(-1,-1),8.5),
            ("TEXTCOLOR",(0,0),(0,-1),MG),
            ("BOTTOMPADDING",(0,0),(-1,-1),3),
            ("TOPPADDING",(0,0),(-1,-1),1),
        ]))
        story.append(ct2)

        # Footer line
        story.append(Spacer(1, 0.35*inch))
        story.append(HRFlowable(width="100%",thickness=0.5,color=BG2))
        story.append(Spacer(1,3))
        story.append(Paragraph(f"\u00a9 {datetime.now().year} {display_brand_name}", ssm))
        comp_photo_entries = []
        for c in vc:
            comp_pdf_photo = build_pdf_image(c.get("photo"), 3.0*inch, max_height=2.15*inch)
            if comp_pdf_photo is not None:
                comp_photo_entries.append((comp_pdf_photo, sanitize_pdf_text(f"Comp {c['n']} - {c['addr'] or 'Comparable'}")))
        swot_text_style = ParagraphStyle("swot_text", parent=sb, fontSize=8.5, leading=11, alignment=TA_JUSTIFY)

        def swot_box(title, body):
            box = Table([
                [Paragraph(f"<b>{title}</b>", ParagraphStyle(f"swot_{title}_head", parent=sb, fontSize=8.5, textColor=white, alignment=TA_CENTER, fontName="Helvetica-Bold"))],
                [Paragraph(sanitize_pdf_text(body or "N/A"), swot_text_style)],
            ], colWidths=[3.65*inch])
            box.setStyle(TableStyle([
                ("BACKGROUND",(0,0),(-1,0),NV),
                ("TEXTCOLOR",(0,0),(-1,0),white),
                ("VALIGN",(0,0),(-1,-1),"TOP"),
                ("TOPPADDING",(0,0),(-1,0),5),
                ("BOTTOMPADDING",(0,0),(-1,0),5),
                ("TOPPADDING",(0,1),(-1,1),8),
                ("BOTTOMPADDING",(0,1),(-1,1),8),
                ("LEFTPADDING",(0,0),(-1,-1),12),
                ("RIGHTPADDING",(0,0),(-1,-1),12),
                ("GRID",(0,0),(-1,-1),table_grid_width,table_grid_color),
            ]))
            return box

        swot_grid = Table([
            [swot_box("Strengths", swot_strengths), swot_box("Weaknesses", swot_weaknesses)],
            [swot_box("Opportunities", swot_opportunities), swot_box("Threats", swot_threats)],
        ], colWidths=[W/2, W/2])
        swot_grid.setStyle(TableStyle([
            ("VALIGN",(0,0),(-1,-1),"TOP"),
            ("LEFTPADDING",(0,0),(-1,-1),0),
            ("RIGHTPADDING",(0,0),(-1,-1),0),
            ("TOPPADDING",(0,0),(-1,-1),4),
            ("BOTTOMPADDING",(0,0),(-1,-1),10),
        ]))

        # â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
        # PAGE 2: PROPERTY SUMMARY + MARKET
        # â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
        # Single-column key-value with clean alignment
        def kv_row(label, val):
            return [
                Paragraph(f"<b>{label}</b>", ParagraphStyle("kl",parent=sb,fontSize=8,textColor=MG,fontName="Helvetica-Bold")),
                Paragraph(sanitize_pdf_text(val), ParagraphStyle("kv",parent=sb,fontSize=9))
            ]

        prop_rows = [
            kv_row("Address", subj_address or "\u2014"),
            kv_row("Folio", subj_folio or "\u2014"),
            kv_row("Property Type", subj_type),
            kv_row("Year Built", str(subj_yr)),
            kv_row("Building Size", f"{subj_sf:,} SF" if subj_sf else "\u2014"),
            kv_row("Lot Size", f"{subj_lot:,} SF ({lot_ac:.3f} acres)" if subj_lot else "\u2014"),
            kv_row("Occupancy", f"{subj_occ:.0f}%"),
            kv_row("Zoning", subj_zone or "\u2014"),
            kv_row("Flood Zone", subj_flood),
            kv_row("Condition", subj_cond),
            kv_row("Current Use", subj_curuse or "\u2014"),
            kv_row("HBU (as improved)", subj_hbu_imp or "\u2014"),
            kv_row("HBU (as vacant)", subj_hbu_vac or "\u2014"),
            kv_row("Owner of Record", st.session_state.get("owner_of_record") or "\u2014"),
            kv_row("Title / Ownership Status", st.session_state.get("title_status") or "\u2014"),
        ]
        pt = Table(prop_rows, colWidths=[1.6*inch, 5.8*inch])
        pt.setStyle(TableStyle([
            ("ROWBACKGROUNDS",(0,0),(-1,-1),[white,LG]),
            ("TOPPADDING",(0,0),(-1,-1),5),
            ("BOTTOMPADDING",(0,0),(-1,-1),5),
            ("LINEBELOW",(0,0),(-1,-1),0.3,BG2),
        ]))
        swot_has_content = any([swot_strengths, swot_weaknesses, swot_opportunities, swot_threats])
        if swot_has_content:
            qualitative_flowables = [
                band("QUALITATIVE MARKET ANALYSIS (SWOT)"),
                Spacer(1,8),
                swot_grid,
                Spacer(1,8),
                band("PROPERTY SUMMARY"),
                Spacer(1,6),
                pt,
            ]
        else:
            qualitative_flowables = [
                band("PROPERTY SUMMARY"),
                Spacer(1,6),
                pt,
            ]

        if subj_desc:
            qualitative_flowables.extend([Spacer(1,6), Paragraph(f"<b>Description:</b> {sanitize_pdf_text(subj_desc)}", ParagraphStyle("desc", parent=sb, leading=14))])
        pdf_ext_assumptions = st.session_state.get("extraordinary_assumptions") or ""
        if pdf_ext_assumptions:
            qualitative_flowables.append(Spacer(1,8))
            qualitative_flowables.append(Paragraph(f"<b>Extraordinary Assumptions and Hypothetical Conditions:</b> {sanitize_pdf_text(pdf_ext_assumptions)}", ParagraphStyle("ext_assump", parent=sb, leading=13, fontSize=8.5)))
        if has_env and env_desc:
            qualitative_flowables.append(Spacer(1,10))
            qualitative_flowables.append(Paragraph(f"<font color='#C0392B'><b>Environmental:</b> {sanitize_pdf_text(env_desc)}</font>", sb))
        if has_code and code_desc:
            qualitative_flowables.append(Spacer(1,10))
            qualitative_flowables.append(Paragraph(f"<font color='#C0392B'><b>Code Violations:</b> {sanitize_pdf_text(code_desc)}</font>", sb))

        market_flowables = [Spacer(1,8), band("MARKET OVERVIEW"), Spacer(1,4)]
        if mkt_factors:
            market_flowables.append(Paragraph(sanitize_pdf_text(mkt_factors), sb))
            market_flowables.append(Spacer(1,10))
        mki = []
        if mkt_vac > 0: mki.append(f"Submarket Vacancy: {mkt_vac:.1f}%")
        if mkt_abs: mki.append(f"Absorption: {mkt_abs}")
        if mki:
            market_flowables.append(Paragraph(sanitize_pdf_text(" | ".join(mki)), ParagraphStyle("mk",parent=sb,textColor=MG,fontSize=8.5)))
        elif not mkt_factors:
            market_flowables.append(Paragraph("<i>Market data to be supplemented.</i>", ParagraphStyle("mi",parent=sb,textColor=MG)))

        story.extend(qualitative_flowables)
        if mkt_map_photo is not None:
            map_image = build_pdf_image(mkt_map_photo, 7.4 * inch, max_height=3.5 * inch)
            if map_image is not None:
                market_flowables.append(Spacer(1,10))
                map_wrap = Table([[map_image]], colWidths=[W])
                map_wrap.setStyle(TableStyle([("ALIGN",(0,0),(-1,-1),"CENTER"),("LEFTPADDING",(0,0),(-1,-1),0),("RIGHTPADDING",(0,0),(-1,-1),0)]))
                market_flowables.append(map_wrap)
                market_flowables.append(Spacer(1,10))
        if any(v > 0 for v in [mkt_pop_1m, mkt_pop_3m, mkt_pop_5m, mkt_med_inc_1m, mkt_med_inc_3m, mkt_med_inc_5m, mkt_vpd_1m, mkt_vpd_3m, mkt_vpd_5m]):
            market_flowables.append(Spacer(1,10))
            demographics_table = Table([
                ["Radius", "1-Mile", "3-Mile", "5-Mile"],
                ["Population", f"{int(mkt_pop_1m):,}", f"{int(mkt_pop_3m):,}", f"{int(mkt_pop_5m):,}"],
                ["Median Income", f"${int(mkt_med_inc_1m):,}", f"${int(mkt_med_inc_3m):,}", f"${int(mkt_med_inc_5m):,}"],
                ["Traffic (VPD)", f"{int(mkt_vpd_1m):,}", f"{int(mkt_vpd_3m):,}", f"{int(mkt_vpd_5m):,}"],
            ], colWidths=[1.85*inch, 1.85*inch, 1.85*inch, 1.85*inch], hAlign='CENTER')
            demographics_table.setStyle(TableStyle([
                ("BACKGROUND",(0,0),(-1,0),NV),
                ("TEXTCOLOR",(0,0),(-1,0),white),
                ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
                ("ROWBACKGROUNDS",(0,1),(-1,-1),[white,LG]),
                ("ALIGN",(1,0),(-1,-1),"CENTER"),
                ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
                ("GRID",(0,0),(-1,-1),table_grid_width,table_grid_color),
            ]))
            market_flowables.append(demographics_table)
        story.append(KeepTogether(market_flowables))

        # â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
        # PAGE 3: COMPARABLE SALES
        # â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
        story.append(PageBreak())
        if vc:
            # Table header
            metric_header = "$/SF (Land)" if use_land_sales else ("$/Unit" if use_unit_sales else "$/SF")
            ch = ["#","Address","Source","Type","Date","Price","Bldg SF","Lot SF","Yr Blt/Ren",metric_header,"Trend%","Adj%",f"Adj {metric_header}","Wt"]
            cr2 = [ch]
            for c in vc:
                yr_str = str(c["yr"])
                if c["ren"] > 0: yr_str += f" / {c['ren']}"
                comp_flag = "*" if c in flagged_outliers else ""
                lot_str = f'{c["lot"]:,}' if c["lot"] > 0 else "\u2014"
                cr2.append([
                    f"{c['n']}{comp_flag}",
                    Paragraph(sanitize_pdf_text((c["addr"] or "\u2014")[:18]), sbs),
                    Paragraph(sanitize_pdf_text((c["source"] or "\u2014")[:14]), sbs),
                    Paragraph(sanitize_pdf_text(c["stype"]), sbs),
                    c["dt"].strftime("%m/%y"),
                    fmt_d(c["price"]),
                    f'{c["sf"]:,}',
                    lot_str,
                    yr_str,
                    fmt_d2(c["psf"]),
                    f'{c.get("trend_pct", 0.0):+.1f}%',
                    f'{c["adj"]:+.0f}%',
                    fmt_d2(c["apsf"]),
                    f'{c["wn"]:.0%}'
                ])

            ctbl = Table(cr2, colWidths=[0.20*inch, 0.70*inch, 0.55*inch, 0.80*inch, 0.40*inch, 0.65*inch, 0.55*inch, 0.45*inch, 0.55*inch, 0.60*inch, 0.45*inch, 0.40*inch, 0.60*inch, 0.50*inch])
            ctbl.setStyle(TableStyle([
                ("BACKGROUND",(0,0),(-1,0),NV),
                ("TEXTCOLOR",(0,0),(-1,0),white),
                ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
                ("FONTSIZE",(0,0),(-1,0),7),
                ("FONTSIZE",(0,1),(-1,-1),7.5),
                ("ALIGN",(0,0),(0,-1),"CENTER"),
                ("ALIGN",(5,0),(-1,-1),"RIGHT"),
                ("ROWBACKGROUNDS",(0,1),(-1,-1),[white,LG]),
                ("TOPPADDING",(0,0),(-1,0),comp_header_padding),
                ("BOTTOMPADDING",(0,0),(-1,0),comp_header_padding),
                ("TOPPADDING",(0,1),(-1,-1),comp_body_padding),
                ("BOTTOMPADDING",(0,1),(-1,-1),comp_body_padding),
                ("GRID",(0,0),(-1,-1),table_grid_width,table_grid_color),
            ]))
            story.append(KeepTogether([band("COMPARABLE SALES ANALYSIS"), Spacer(1,8)]))
            story.append(ctbl)
            if flagged_outliers:
                audit_rows = [[Paragraph("<b>STATISTICAL AUDIT & OUTLIER DISCLOSURE</b>", ParagraphStyle("audit_head", parent=sb, fontSize=9, fontName="Helvetica-Bold", textColor=HexColor("#991B1B")))]]
                for c in flagged_outliers:
                    variance_pct = (abs(c["apsf"] - pmed) / pmed * 100) if pmed > 0 else 0.0
                    audit_rows.append([Paragraph(
                        f"Comp {c['n']} reflects a {variance_pct:.1f}% variance from the adjusted median benchmark based on adjusted {metric_header}.",
                        ParagraphStyle("audit_body", parent=sbs, textColor=HexColor("#7F1D1D"))
                    )])
                audit_table = Table(audit_rows, colWidths=[W])
                audit_table.setStyle(TableStyle([
                    ("BACKGROUND",(0,0),(-1,-1),HexColor("#FFF1F2")),
                    ("BOX",(0,0),(-1,-1),1,HexColor("#FCA5A5")),
                    ("TOPPADDING",(0,0),(-1,-1),6),
                    ("BOTTOMPADDING",(0,0),(-1,-1),6),
                    ("LEFTPADDING",(0,0),(-1,-1),10),
                    ("RIGHTPADDING",(0,0),(-1,-1),10),
                ]))
                story.append(Spacer(1,6))
                story.append(audit_table)
            if flagged_outliers:
                story.append(Spacer(1,4))
                story.append(Paragraph("*Asterisk denotes a comparable identified as a statistical outlier based on adjusted market metric variance greater than 15% from the adjusted median.", ssm))
            if manual_weight_override:
                story.append(Spacer(1,4))
                story.append(Paragraph("*Weights manually assigned by Broker based on qualitative market factors and asset-specific similarities.", ssm))

            # Adjustment notes - with spacing
            has_notes = any(c["notes"] for c in vc)
            if has_notes:
                story.append(Spacer(1,8))
                story.append(Paragraph("<b>Adjustment Notes</b>", ParagraphStyle("anh",parent=sb,fontSize=8.5,fontName="Helvetica-Bold",textColor=NV)))
                story.append(Spacer(1,2))
                for c in vc:
                    if c["notes"]:
                        story.append(Paragraph(f"\u2022 Comp {c['n']}: {sanitize_pdf_text(c['notes'])}", ParagraphStyle("an",parent=sb,fontSize=8,textColor=MG,leftIndent=6,spaceBefore=1)))
                story.append(Spacer(1,6))

            # Value summary line
            story.append(Spacer(1,6))
            val_line = Table([[
                Paragraph(f"Weighted Adjusted {metric_header}: <b>{fmt_d2(weighted_psf)}</b>", ParagraphStyle("wl",parent=sb,fontSize=9.5,textColor=NV)),
                Paragraph(f"<b>Sales Comparison Value: {fmt_d(sales_comp_value)}</b>", ParagraphStyle("sv",parent=sb,fontSize=9.5,textColor=NV,alignment=TA_RIGHT,fontName="Helvetica-Bold"))
            ]], colWidths=[W/2, W/2])
            val_line.setStyle(TableStyle([
                ("BACKGROUND",(0,0),(-1,-1),LG),
                ("TOPPADDING",(0,0),(-1,-1),6),
                ("BOTTOMPADDING",(0,0),(-1,-1),6),
                ("LEFTPADDING",(0,0),(0,0),10),
                ("RIGHTPADDING",(-1,-1),(-1,-1),10),
                ("BOX",(0,0),(-1,-1),0.5,BG2),
            ]))
            story.append(val_line)

            # Chart
            story.append(Spacer(1,8))
            cbuf = io.BytesIO()
            fig, ax = plt.subplots(figsize=(6, 1.8))
            fig.patch.set_facecolor("white"); ax.set_facecolor("white")
            lbs = [f"C{c['n']}" for c in vc] + ["Subject"]
            vls = [c["apsf"] for c in vc] + [weighted_psf]
            ax.bar(lbs, vls, color=["#002147"]*len(vc)+["#94A3B8"], width=0.5)
            ax.set_ylim(0, max(vls)*1.22)
            for ix, v2 in enumerate(vls):
                ax.text(ix, v2+max(vls)*0.025, fmt_d(v2), ha="center", fontsize=7.5, fontweight="bold")
            chart_metric_label = "Adj $/SF (Land)" if use_land_sales else ("Adj $/Unit" if use_unit_sales else "Adj $/SF (Bldg)")
            ax.set_ylabel(chart_metric_label, fontsize=7, color="#64748B")
            ax.tick_params(labelsize=7)
            ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False)
            plt.tight_layout()
            fig.savefig(cbuf, format="png", dpi=180, bbox_inches="tight"); plt.close(fig); cbuf.seek(0)
            story.append(Image(cbuf, width=5.5*inch, height=1.6*inch))
            if comp_photo_entries:
                story.append(Spacer(1,8))
                story.append(KeepTogether([band("PHOTOGRAPHIC EXHIBITS"), Spacer(1,8)]))
                gallery_rows = []
                for idx in range(0, len(comp_photo_entries), 2):
                    pair = comp_photo_entries[idx:idx+2]
                    image_row = []
                    caption_row = []
                    for photo_img, caption in pair:
                        image_row.append(photo_img)
                        caption_row.append(Paragraph(caption, ParagraphStyle("pcap", parent=sbs, alignment=TA_CENTER, textColor=MG)))
                    while len(image_row) < 2:
                        image_row.append(Spacer(1, 0.1*inch))
                        caption_row.append(Paragraph("", ParagraphStyle("pcap_empty", parent=sbs, alignment=TA_CENTER, textColor=MG)))
                    gallery_rows.append(image_row)
                    gallery_rows.append(caption_row)
                photo_gallery = Table(gallery_rows, colWidths=[W/2, W/2])
                photo_gallery.setStyle(TableStyle([
                    ("ALIGN",(0,0),(-1,-1),"CENTER"),
                    ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
                    ("TOPPADDING",(0,0),(-1,-1),6),
                    ("BOTTOMPADDING",(0,0),(-1,-1),6),
                ]))
                story.append(photo_gallery)
        else:
            story.append(KeepTogether([band("COMPARABLE SALES ANALYSIS"), Spacer(1,8), Paragraph("<i>No comparable sales data entered.</i>", ParagraphStyle("nc",parent=sb,textColor=MG))]))

        # â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
        # PAGE 4: INCOME (if used) + RECONCILIATION
        # â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
        has_inc = income_value > 0

        if has_inc:
            story.append(PageBreak())
            income_band = band("INCOME ANALYSIS")
            unit_mix_has_rows = False

            if use_unit_mix and not unit_mix_summary.empty:
                unit_mix_heading = Paragraph("<b>UNIT MIX & RENT ROLL SUMMARY</b>", ssec)
                unit_mix_pdf = [["Unit Type", "Count", "Vacant", "Total SF", "Monthly Rent", "Market Rent"]]
                for _, row in unit_mix_summary.iterrows():
                    if row["Count"] > 0 or row["Monthly Rent"] > 0 or row["Market Rent"] > 0:
                        unit_mix_pdf.append([
                            sanitize_pdf_text(str(row["Unit Type"] or "\u2014")),
                            f'{int(row["Count"]):,}' if row["Count"] else "0",
                            f'{int(row["Vacant Count"]):,}' if row["Vacant Count"] else "0",
                            f'{int(row["Total SF"]):,}' if row["Total SF"] else "0",
                            fmt_d(row["Monthly Rent"]),
                            fmt_d(row["Market Rent"]),
                        ])
                if len(unit_mix_pdf) > 1:
                    unit_mix_pdf.append([
                        "TOTAL",
                        f'{int(unit_mix_summary["Count"].sum()):,}',
                        f'{int(unit_mix_summary["Vacant Count"].sum()):,}',
                        f'{int(unit_mix_summary["Total SF"].sum()):,}',
                        fmt_d(unit_mix_summary["Monthly Rent"].sum()),
                        fmt_d(unit_mix_summary["Market Rent"].sum()),
                    ])
                if len(unit_mix_pdf) > 1:
                    umt = Table(unit_mix_pdf, colWidths=[1.6 * inch, 0.7 * inch, 0.7 * inch, 0.9 * inch, 1.7 * inch, 1.8 * inch])
                    umt.setStyle(TableStyle([
                        ("BACKGROUND",(0,0),(-1,0),NV),
                        ("TEXTCOLOR",(0,0),(-1,0),white),
                        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
                        ("FONTNAME",(0,-1),(-1,-1),"Helvetica-Bold"),
                        ("FONTSIZE",(0,0),(-1,-1),8),
                        ("ALIGN",(1,1),(-1,-1),"RIGHT"),
                        ("TOPPADDING",(0,0),(-1,-1),3),
                        ("BOTTOMPADDING",(0,0),(-1,-1),3),
                        ("GRID",(0,0),(-1,-1),table_grid_width,table_grid_color),
                        ("ROWBACKGROUNDS",(0,1),(-1,-1),[white,LG]),
                    ]))
                    story.append(KeepTogether([income_band, Spacer(1,6), unit_mix_heading, umt]))
                    story.append(Spacer(1,6))
                    unit_mix_has_rows = True

            income_tax_label = "Pro-Forma Ad Valorem Taxes (Reassessed)" if use_tax_reassessment else "Property Taxes"
            idat = [["", "Amount"]]
            if use_unit_mix:
                idat.extend([
                    ["Gross Potential Rent", fmt_d(gpr)],
                    ["Less: Physical Vacancy", f"({fmt_d(physical_vacancy_loss)})"],
                    ["Less: Loss to Lease (Current vs. Market)", f"({fmt_d(loss_to_lease)})"],
                    ["Plus: Expense Reimbursements", fmt_d(reimb)],
                    [f"Less: Credit Loss ({vac_rate:.0f}%)", f"({fmt_d(credit_loss_amount)})"],
                ])
            else:
                idat.extend([
                    ["Gross Potential Rent", fmt_d(gpr)],
                    ["Plus: Expense Reimbursements", fmt_d(reimb)],
                    [f"Less: Vacancy & Credit Loss ({vac_rate:.0f}%)", f"({fmt_d(credit_loss_amount)})"],
                ])
            idat.extend([
                ["Effective Gross Income", fmt_d(egi)],
                [income_tax_label, f"({fmt_d(income_tax_line)})"],
                ["Insurance", f"({fmt_d(opex_ins)})"],
                [f"Management Fee ({opex_mgmt_pct:.1f}%)", f"({fmt_d(opex_mgmt_d)})"],
                ["Maintenance", f"({fmt_d(opex_maint)})"],
                ["Utilities", f"({fmt_d(opex_util)})"],
                ["Reserves", f"({fmt_d(opex_res)})"],
                ["Other Expenses", f"({fmt_d(opex_oth)})"],
                ["Total Operating Expenses", f"({fmt_d(total_opex)})"],
                ["Net Operating Income (NOI)", fmt_d(noi)],
            ])
            if use_unit_mix and loss_to_lease > 0:
                idat.append(["Value-Add Upside: Mark-to-Market Rent", fmt_d(loss_to_lease)])
            it = Table(idat, colWidths=[4.2*inch, 3.2*inch])
            income_style = [
                ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
                ("FONTSIZE",(0,0),(-1,-1),9),
                ("BACKGROUND",(0,0),(-1,0),NV),
                ("TEXTCOLOR",(0,0),(-1,0),white),
                ("ALIGN",(1,0),(1,-1),"RIGHT"),
                ("GRID",(0,0),(-1,-1),table_grid_width,table_grid_color),
                ("ROWBACKGROUNDS",(0,1),(-1,-1),[white,LG]),
                ("TOPPADDING",(0,0),(-1,-1),3),
                ("BOTTOMPADDING",(0,0),(-1,-1),3),
            ]
            income_row_map = {row[0]: idx for idx, row in enumerate(idat)}
            for key_label in ["Effective Gross Income", "Total Operating Expenses", "Net Operating Income (NOI)", "Value-Add Upside: Mark-to-Market Rent"]:
                idx = income_row_map.get(key_label)
                if idx is not None:
                    income_style.append(("FONTNAME",(0,idx),(-1,idx),"Helvetica-Bold"))
            for navy_label in ["Total Operating Expenses", "Net Operating Income (NOI)", "Value-Add Upside: Mark-to-Market Rent"]:
                navy_idx = income_row_map.get(navy_label)
                if navy_idx is not None:
                    income_style.append(("BACKGROUND",(0,navy_idx),(-1,navy_idx),NV))
                    income_style.append(("TEXTCOLOR",(0,navy_idx),(-1,navy_idx),white))
            it.setStyle(TableStyle(income_style))
            story.append(KeepTogether([income_band, Spacer(1,6), it]) if not unit_mix_has_rows else it)
            story.append(Spacer(1,15))
            if use_tax_reassessment:
                story.append(Spacer(1,4))
                story.append(Paragraph(f"Taxes estimated at {tax_rate_pct:.2f}% of opinion of value per Miami-Dade post-acquisition standards.", ssm))
            story.append(Spacer(1,6))
            iv_style = ParagraphStyle("iv_inline", parent=sb, fontSize=9.5, leading=14, textColor=NV)
            story.append(Paragraph(f"Market Cap Rate: <b>{inc_cap:.2f}%</b>  |  Tenant Risk Adj: <b>+{risk_premium:.2f}%</b>  |  <b>Adj. Income Value: {fmt_d(income_value)}</b>", iv_style))
            pdf_dscr = st.session_state.get("dscr")
            pdf_ads = st.session_state.get("annual_debt_service")
            pdf_loan_amount = st.session_state.get("loan_amount")
            pdf_irr = st.session_state.get("irr_value")
            pdf_breakeven_occ = st.session_state.get("breakeven_occ")
            pdf_loan_ltv = st.session_state.get("loan_ltv")
            pdf_loan_rate = st.session_state.get("loan_rate")
            if pdf_dscr is not None and pdf_ads is not None and pdf_loan_amount is not None and pdf_loan_ltv is not None and pdf_loan_rate is not None:
                story.append(Spacer(1,4))
                story.append(Paragraph(f"DSCR: <b>{pdf_dscr:.2f}x</b>  |  Annual Debt Service: <b>{fmt_d(pdf_ads)}</b>  |  Loan Amount: <b>{fmt_d(pdf_loan_amount)}</b>  |  Assumed LTV: <b>{pdf_loan_ltv:.0f}%</b> on BOV Value  |  Rate: <b>{pdf_loan_rate:.2f}%</b>", iv_style))
            if pdf_breakeven_occ is not None:
                story.append(Spacer(1,4))
                if pdf_breakeven_occ > 100.0:
                    story.append(Paragraph(
                        f"Break-Even Occupancy (1.20x DSCR): <b>Unachievable at {pdf_breakeven_occ:.1f}% of GPR</b> — NOI at full occupancy is insufficient to cover modeled debt service.",
                        ParagraphStyle("beo_pdf_alert", parent=sb, fontSize=9.5, leading=14, textColor=HexColor("#DC2626"))
                    ))
                else:
                    pdf_subj_occ = st.session_state.get("subj_occ", 0.0)
                    beo_suffix = f" (Current occupancy: {pdf_subj_occ:.0f}%)" if pdf_subj_occ > 0 else " (Subject occupancy not entered)"
                    story.append(Paragraph(
                        f"Break-Even Occupancy (1.20x DSCR): <b>{pdf_breakeven_occ:.1f}%</b>{beo_suffix}",
                        iv_style
                    ))
            if pdf_irr is not None:
                story.append(Spacer(1,4))
                story.append(Paragraph(f"All-Cash Return (Unlevered, Pre-Tax IRR, Net of Exit Costs): <b>{pdf_irr:.2f}%</b>  |  Discount Rate: <b>{st.session_state.get('ddr', 8.5):.2f}%</b>  |  Exit Selling Cost: <b>{st.session_state.get('exit_sc', 4.0):.1f}%</b>", iv_style))
            story.append(Spacer(1,3))
            story.append(Paragraph("*Cap Rate adjusted to reflect the increased risk associated with the current tenant credit profile and local market vacancy risks.", ssm))
            if stab_value > 0:
                story.append(Spacer(1,6))
                story.append(Paragraph(f"Stabilized Value: <b>{fmt_d(stab_value)}</b>", iv_style))
            if st.session_state.get("use_dcf", False) and dcf_value > 0:
                story.append(Spacer(1,6))
                story.append(Paragraph(f"Discounted Cash Flow (NPV, Gross Reversion — pre-selling-costs): <b>{fmt_d(dcf_value)}</b>", iv_style))
            story.append(Spacer(1,20))

        if use_cost:
            cost_dat = [["Land Value Estimate", fmt_d(cost_land)]]
            if subj_type != "Land":
                cost_dat.extend([
                    ["Total Replacement Cost", fmt_d(tot_repl)],
                    ["Total Depreciation Percentage", f"{td:.0f}%"],
                ])
            cost_dat.append(["Cost Approach Value", fmt_d(cost_value)])
            cost_table = Table(cost_dat, colWidths=[4.4*inch, 3.0*inch])
            cost_table.setStyle(TableStyle([
                ("GRID",(0,0),(-1,-1),table_grid_width,table_grid_color),
                ("ROWBACKGROUNDS",(0,0),(-1,-1),[white,LG]),
                ("FONTNAME",(0,0),(0,-1),"Helvetica-Bold"),
                ("FONTNAME",(0,-1),(-1,-1),"Helvetica-Bold"),
                ("ALIGN",(1,0),(1,-1),"RIGHT"),
                ("TOPPADDING",(0,0),(-1,-1),3),
                ("BOTTOMPADDING",(0,0),(-1,-1),3),
            ]))
            story.append(KeepTogether([band("COST APPROACH"), Spacer(1,6), cost_table, Spacer(1,10)]))

        # RECONCILIATION
        if not has_inc:
            story.append(PageBreak())

        recon_band = band("RECONCILIATION & RECOMMENDATION")

        wt_s = wts.get("Sales Comparison",0)
        wt_i = wts.get("Income Capitalization",0)
        wt_c = wts.get("Cost Approach",0)
        pdf_capex = st.session_state.get("capex_deduct", 0.0)
        as_is_value = recon_val - pdf_capex

        rdat = [["Valuation Approach","Indicated Value","Weight","Weighted Value"]]
        if sales_comp_value > 0: rdat.append(["Sales Comparison", fmt_d(sales_comp_value), f"{wt_s:.0f}%", fmt_d(sales_comp_value*wt_s/100)])
        if income_value > 0: rdat.append(["Income Capitalization", fmt_d(income_value), f"{wt_i:.0f}%", fmt_d(income_value*wt_i/100)])
        if cost_value > 0: rdat.append(["Cost Approach", fmt_d(cost_value), f"{wt_c:.0f}%", fmt_d(cost_value*wt_c/100)])
        rdat.append(["Less: Deferred Maintenance/CapEx", "", "", f"-{fmt_d(pdf_capex)}"])
        rdat.append(["RECONCILED VALUE","","100%", fmt_d(as_is_value)])

        rt = Table(rdat, colWidths=[2.2*inch, 1.5*inch, 0.8*inch, 1.5*inch])
        rt.setStyle(TableStyle([
            ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
            ("FONTSIZE",(0,0),(-1,-1),9),
            ("BACKGROUND",(0,0),(-1,0),NV),
            ("TEXTCOLOR",(0,0),(-1,0),white),
            ("ROWBACKGROUNDS",(0,1),(-1,-1),[white,LG]),
            ("FONTNAME",(0,-1),(-1,-1),"Helvetica-Bold"),
            ("BACKGROUND",(0,-1),(-1,-1),LG),
            ("TEXTCOLOR",(-1,-1),(-1,-1),NV),
            ("ALIGN",(1,0),(-1,-1),"RIGHT"),
            ("TOPPADDING",(0,0),(-1,-1),summary_padding),
            ("BOTTOMPADDING",(0,0),(-1,-1),summary_padding),
            ("GRID",(0,0),(-1,-1),table_grid_width,table_grid_color),
        ]))
        story.append(KeepTogether([recon_band, Spacer(1,8), rt]))
        story.append(Spacer(1, 8))
        story.append(Paragraph(f"<b>LIQUIDATION VALUE ({liquidation_discount:.0f}% Discount from As-Is):</b> <font color='#DC2626'>{fmt_d(liquidation_value)}</font>", ParagraphStyle("olv_pdf", parent=sb, textColor=NV, fontSize=9.5)))
        if recon_narrative:
            story.append(Spacer(1,8))
            story.append(Paragraph("Valuation Reconciliation", ssec))
            story.append(Paragraph(sanitize_pdf_text(recon_narrative), sb))
        if sync_adjusted_cap > 0 and sensitivity_matrix.size:
            occ = sync_subj_occ
            sm_head = Paragraph("<b>MARKET SENSITIVITY & STRESS TEST</b>", ssec)
            sm_dat = [[
                "",
                f"{sensitivity_occ_steps[0]:.0f}% Collection",
                f"{sensitivity_occ_steps[1]:.0f}% Collection",
                f"{sensitivity_occ_steps[2]:.0f}% Collection",
            ]]
            for idx, cap_step in enumerate(sensitivity_cap_steps):
                sm_dat.append([
                    f"{cap_step:.2f}%",
                    fmt_d(sensitivity_matrix[idx, 0]),
                    fmt_d(sensitivity_matrix[idx, 1]),
                    fmt_d(sensitivity_matrix[idx, 2]),
                ])
            smt = Table(sm_dat, colWidths=[1.1*inch, 1.6*inch, 1.6*inch, 1.6*inch])
            smt.setStyle(TableStyle([
                ("BACKGROUND",(0,0),(-1,0),NV),
                ("TEXTCOLOR",(0,0),(-1,0),white),
                ("BACKGROUND",(0,0),(0,-1),NV),
                ("TEXTCOLOR",(0,0),(0,-1),white),
                ("BACKGROUND",(2,2),(2,2),LG),
                ("TEXTCOLOR",(3,1),(3,1),HexColor("#16A34A")),
                ("TEXTCOLOR",(1,3),(1,3),HexColor("#DC2626")),
                ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
                ("FONTNAME",(0,0),(0,-1),"Helvetica-Bold"),
                ("FONTSIZE",(0,0),(-1,-1),8),
                ("ALIGN",(1,1),(-1,-1),"RIGHT"),
                ("ALIGN",(0,0),(0,-1),"RIGHT"),
                ("TOPPADDING",(0,0),(-1,-1),5),
                ("BOTTOMPADDING",(0,0),(-1,-1),5),
                ("GRID",(0,0),(-1,-1),table_grid_width,table_grid_color),
            ]))
            stress_block = KeepTogether([
                Spacer(1,8),
                sm_head,
                Spacer(1,4),
                smt,
                Spacer(1,4),
                Paragraph("Note: Matrix represents +/- 50 bps cap rate volatility and +/- 5% collection rate variation from the current effective rate. Red values indicate high-risk downside scenarios.", ssm),
            ])
            story.append(stress_block)
        story.append(Spacer(1,6))
        story.append(Paragraph(f"<b>Data Reliability:</b> This valuation has a <b>{st.session_state.get('conf_score', 'N/A')}</b> statistical confidence score, with a Coefficient of Variation (COV) of <b>{st.session_state.get('cov', 0):.1f}%</b>.", sb))
        story.append(Spacer(1,3))
        story.append(Paragraph("*A lower COV indicates higher data consistency and tighter market clustering among selected comparable sales.", ssm))
        story.append(Spacer(1,12))

        recommendation_box_style = ParagraphStyle("rpbox",parent=sb,alignment=TA_CENTER,leading=12,textColor=NV,fontName="Helvetica-Bold")
        rec_price_style = ParagraphStyle("rp_price", parent=recommendation_box_style, leading=28)
        rpb_inner = Table([
            [Paragraph("<font size='8' color='#64748B'>RECOMMENDED LIST PRICE</font>", recommendation_box_style)],
            [Paragraph(f"<font size='22'>{fmt_d(rec_price)}</font>", rec_price_style)],
        ], colWidths=[4.5*inch])
        rpb_inner.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,-1),LG),
            ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
            ("ALIGN",(0,0),(-1,-1),"CENTER"),
            ("BOX",(0,0),(-1,-1),1,SV),
            ("TOPPADDING",(0,0),(-1,-1),4),
            ("BOTTOMPADDING",(0,-1),(-1,-1),16),
        ]))
        rpb_wrap = Table([[rpb_inner]], colWidths=[W])
        rpb_wrap.setStyle(TableStyle([("ALIGN",(0,0),(-1,-1),"CENTER")]))
        story.append(rpb_wrap)

        story.append(Spacer(1,4))
        _pdf_cov = st.session_state.get("cov", 0.0)
        _pdf_band = 0.05 if _pdf_cov < 10.0 else (0.10 if _pdf_cov <= 15.0 else 0.15)
        if st.session_state.get("conf_score", "N/A") != "N/A":
            story.append(Paragraph(f"Price Range: {fmt_d(rec_price*(1-_pdf_band))} \u2013 {fmt_d(rec_price*(1+_pdf_band))}  |  Estimated Marketing Time: {rec_mkt} ({rec_dom} DOM)  |  Range reflects \u00b1{_pdf_band:.0%} based on {st.session_state.get('conf_score','N/A')} statistical confidence (COV: {_pdf_cov:.1f}%)", ParagraphStyle("rng",parent=sb,fontSize=8.5,textColor=MG,alignment=TA_CENTER)))
        else:
            story.append(Paragraph(f"Estimated Marketing Time: {rec_mkt} ({rec_dom} DOM)", ParagraphStyle("rng",parent=sb,fontSize=8.5,textColor=MG,alignment=TA_CENTER)))
        if rec_just:
            story.append(Paragraph(f"<b>Pricing Justification:</b> <i>{sanitize_pdf_text(rec_just)}</i>", ParagraphStyle("rji",parent=sb,textColor=MG,alignment=TA_CENTER,fontSize=8.5)))

        if rec_narr:
            story.append(Spacer(1,10))
            story.append(Paragraph("<b>Disposition Strategy</b>", ssec))
            story.append(Paragraph(sanitize_pdf_text(rec_narr), sb))
        # â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
        # LAST PAGE: CONDITIONS + CERT + SIGNATURES
        # â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
        story.append(PageBreak())
        if not inst_mode:
            story.append(band("ABOUT ESTUPINAN GROUP"))
            story.append(Spacer(1,6))
            story.append(Paragraph(
                "Estupinan Group is a premier commercial advisory team at First Service Realty by ERA, specializing in high-value asset valuation, disposition or transfer, and capital recovery strategies across the Miami-Dade market. Our institutional-grade analytics provide lenders and private equity clients with the data-driven certainty required for complex real estate transactions.",
                sb
            ))
            story.append(Spacer(1,10))
        statutory_disclosure = Table([[
            Paragraph(
                "<b>This opinion of value was prepared by a licensed real estate broker or sales associate and is not a certified appraisal. This opinion may not be used for purposes of a federally related transaction as defined in the Financial Institutions Reform, Recovery, and Enforcement Act of 1989 (FIRREA).</b>",
                ParagraphStyle("stat_box", parent=sb, fontSize=8.2, leading=11, textColor=HexColor("#7F1D1D"))
            )
        ]], colWidths=[W])
        statutory_disclosure.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,-1),HexColor("#FFF1F2")),
            ("BOX",(0,0),(-1,-1),1.5,HexColor("#DC2626")),
            ("TOPPADDING",(0,0),(-1,-1),8),
            ("BOTTOMPADDING",(0,0),(-1,-1),8),
            ("LEFTPADDING",(0,0),(-1,-1),10),
            ("RIGHTPADDING",(0,0),(-1,-1),10),
        ]))
        story.append(statutory_disclosure)
        story.append(Spacer(1,8))

        for lm in [
            "This BOV is not a certified appraisal and should not be relied upon as such.",
            "The value stated is an opinion based on information believed reliable but not guaranteed.",
            "The broker assumes no responsibility for legal matters including title, encumbrances, or liens.",
            "This opinion is subject to change based on new information or market conditions.",
            "No responsibility is assumed for accuracy of information furnished by third parties.",
            "This opinion is prepared by a licensed Florida real estate broker or sales associate in compliance with Florida Statute \u00a7475.01 et seq. governing real estate brokerage practice. This report is not a certified appraisal and is not subject to the Uniform Standards of Professional Appraisal Practice (USPAP).",
        ]:
            story.append(Paragraph(f"\u2022 {lm}", ParagraphStyle("lm",parent=sb,leftIndent=8,spaceBefore=1.5,fontSize=8,leading=10.5)))

        story.append(Spacer(1,10))
        story.append(Paragraph("<b>Scope of Work and Intended Use</b>", ssec))
        story.append(Paragraph(
            f"Intended Use: <b>{sanitize_pdf_text(pdf_intended_use)}</b>  |  Intended User: <b>{sanitize_pdf_text(pdf_intended_user)}</b>.",
            sb
        ))
        story.append(Paragraph(
            "This BOV is prepared exclusively for the stated intended use and intended user, and reliance by any other party or for any other purpose is expressly disclaimed.",
            sb
        ))
        story.append(Spacer(1,10))
        story.append(Paragraph("<b>CERTIFICATION</b>", ssec))
        cert_subject = f"the property located at {sanitize_pdf_text(subj_address.strip())}" if subj_address else "the property that is the subject of this report"
        story.append(Paragraph(
            f"We certify that, to the best of our knowledge and belief: the statements of fact contained "
            f"in this report are true and correct; the reported analyses, opinions, and conclusions are "
            f"limited only by the reported assumptions and limiting conditions; we have no present or "
            f"prospective interest in {cert_subject}; we have no bias "
            f"with respect to the property or parties involved; our compensation is not contingent upon "
            f"reporting a predetermined value; our analyses, opinions, and conclusions were developed "
            f"in conformity with the ethical obligations applicable to licensed Florida real estate professionals under Florida Statute \u00a7475 and the rules of the Florida Real Estate Commission.",
            ParagraphStyle("ce",parent=sb,fontSize=8,leading=11)
        ))

        story.append(Spacer(1, 0.4*inch))
        story.append(HRFlowable(width="100%",thickness=1,color=NV))
        story.append(Spacer(1, 6))

        story.append(expert_signature_table())

        story.append(Spacer(1, 14))
        story.append(HRFlowable(width="100%",thickness=0.5,color=BG2))
        story.append(Paragraph(f"\u00a9 {datetime.now().year} {display_brand_name}", ssm))

        doc = SimpleDocTemplate(buf, pagesize=letter, leftMargin=0.55*inch, rightMargin=0.55*inch, topMargin=0.3*inch, bottomMargin=0.3*inch, title="Broker Opinion of Value", author=display_brand_name)
        doc.build(story, onFirstPage=draw_page_chrome, onLaterPages=draw_page_chrome)
        buf.seek(0)
        return buf

    g1,g2,g3 = st.columns([1,1.5,1])
    with g2:
        if st.button("Generate BOV PDF", use_container_width=True, type="primary", disabled=pdf_button_disabled):
            if not subj_address: st.error("Enter a property address in Step 1.")
            elif rec_price == 0: st.error("Enter a recommended list price.")
            else:
                if not prep_for:
                    st.warning("Prepared For is blank. The PDF cover page will show a generic placeholder. Consider entering the client or recipient name in Step 1 before submitting.")
                if not intended_user:
                    st.warning("Intended User is blank. This field should identify the specific client or recipient. Consider entering it in Step 1 before submitting.")
                if not rec_narr:
                    st.warning("Broker Recommendation is blank. This field provides the disposition or transfer strategy narrative in the PDF. Consider entering it before generating the report.")
                if not recon_narrative:
                    st.warning("Reconciliation Narrative is blank. This field explains the weighting logic between approaches and is expected in a formal institutional BOV.")
                pdf = gen_pdf()
                ca2 = re.sub(r'[^\w\s-]','',subj_address).strip().replace(' ','_')[:40]
                fn = f"BOV_{ca2}_{datetime.now().strftime('%Y_%m_%d')}.pdf"
                st.download_button(label="Download BOV PDF",data=pdf,file_name=fn,mime="application/pdf",use_container_width=True)
                try:
                    if cloud_db_connected and db_conn is not None:
                        with db_conn.session as session:
                            session.execute(
                                text("INSERT INTO commercial_bov_history (address, ptype, dt, rec_value, sale_price, dom, generated_by) VALUES (:addr, :pt, :dt, :val, :sp, :dom, :auth)"),
                                {
                                    "addr": subj_address,
                                    "pt": subj_type,
                                    "dt": datetime.now().strftime("%Y-%m-%d"),
                                    "val": rec_price,
                                    "sp": last_sale_price,
                                    "dom": rec_dom,
                                    "auth": st.session_state.analyst_name,
                                },
                            )
                            session.commit()
                    elif local_engine is not None:
                        with local_engine.begin() as connection:
                            connection.execute(
                                text("INSERT INTO commercial_bov_history (address, ptype, dt, rec_value, sale_price, dom, generated_by) VALUES (:addr, :pt, :dt, :val, :sp, :dom, :auth)"),
                                {
                                    "addr": subj_address,
                                    "pt": subj_type,
                                    "dt": datetime.now().strftime("%Y-%m-%d"),
                                    "val": rec_price,
                                    "sp": last_sale_price,
                                    "dom": rec_dom,
                                    "auth": st.session_state.analyst_name,
                                },
                            )
                    st.success(f"Saved: **{fn}**")
                except: st.success(f"Generated: **{fn}**")

    st.markdown(sl("BOV History"), unsafe_allow_html=True)
    try:
        if cloud_db_connected and db_conn is not None:
            with db_conn.session as session:
                hist=session.execute(text("SELECT address, ptype, dt, rec_value, sale_price, dom, generated_by FROM commercial_bov_history ORDER BY created DESC LIMIT 20")).mappings().all()
        elif local_engine is not None:
            with local_engine.begin() as connection:
                hist=connection.execute(text("SELECT address, ptype, dt, rec_value, sale_price, dom, generated_by FROM commercial_bov_history ORDER BY created DESC LIMIT 20")).mappings().all()
        else:
            hist=[]
        if hist:
            for r in hist:
                ad,pt2,dt2,vl2,sp2,dm2,generated_by = r["address"],r["ptype"],r["dt"],r["rec_value"],r["sale_price"],r["dom"],r["generated_by"]
                ss2=fmt_d(sp2) if sp2 else "Pending"; dd2=str(dm2) if dm2 else "\u2014"
                st.markdown(f'<div class="mc" style="padding:0.4rem 0.8rem;"><div style="display:flex;justify-content:space-between;font-size:0.8rem;"><span style="color:#D4A843;font-weight:600;">{ad or "\u2014"}</span><span style="color:#6E7681;">{dt2}</span></div><div class="rr"><span class="l">{pt2} | BOV: {fmt_d(vl2)}</span><span class="v">Sale: {ss2} | DOM: {dd2}</span></div><div style="color:{TEXT_MUTED};font-size:0.72rem;">Analyst: {generated_by or "Unknown"}</div></div>', unsafe_allow_html=True)
        else: st.markdown('<div class="co">No BOVs yet.</div>', unsafe_allow_html=True)
    except: st.markdown('<div class="co">Tracker initializes on first report.</div>', unsafe_allow_html=True)

st.markdown(f'<div class="ft">&copy; {datetime.now().year} {BRAND_NAME}</div>', unsafe_allow_html=True)

