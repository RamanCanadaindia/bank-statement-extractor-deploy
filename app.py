from __future__ import annotations

import base64
import importlib
import hmac
import io
import json
import os
import re
import shutil
import sys
import tempfile
import zipfile
from datetime import date
from pathlib import Path

import pandas as pd
import streamlit as st


APP_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(APP_DIR))

import bmo_docling_to_excel as extractor
import financial_statement_generator as fs_generator
import google_sheets_sender as sheets_sender
import mortgage_calculator as mortgage
import transaction_categorizer as categorizer
from real_estate_research_agent import (
    Assumptions as InvestmentAssumptions,
    RentalProvider,
    RealtorSavedSearchProvider,
    SignalProvider,
    ZealtyProvider,
    merge_with_state,
    score_property,
)

importlib.reload(extractor)
importlib.reload(fs_generator)
importlib.reload(sheets_sender)
importlib.reload(mortgage)


st.set_page_config(
    page_title="Raman Financial Services - Accounting Tools",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    :root {
        --rfs-navy: #0b1f3a;
        --rfs-blue: #1268df;
        --rfs-blue-dark: #0b54bd;
        --rfs-soft-blue: #eaf3ff;
        --rfs-bg: #f5f7fb;
        --rfs-border: #dfe5ee;
        --rfs-muted: #64748b;
        --rfs-green: #12805c;
    }
    html, body, [class*="css"] {font-family: Inter, "Segoe UI", Arial, sans-serif;}
    .stApp {background: var(--rfs-bg); color: var(--rfs-navy);}
    [data-testid="stHeader"] {background: rgba(245, 247, 251, 0.92);}
    .block-container {max-width: 1280px; padding: 4.5rem 2rem 4rem;}
    h1, h2, h3 {letter-spacing: 0; color: var(--rfs-navy);}
    h1 {font-size: 1.75rem !important;}
    h2 {font-size: 1.35rem !important;}
    h3 {font-size: 1.05rem !important;}
    .stMarkdown p, label, [data-testid="stCaptionContainer"] {color: #42526a;}

    [data-testid="stSidebar"] {
        background: #ffffff;
        border-right: 1px solid var(--rfs-border);
    }
    [data-testid="stSidebar"] > div:first-child {padding-top: 1rem;}
    [data-testid="stSidebar"] [role="radiogroup"] {gap: 0.2rem;}
    [data-testid="stSidebar"] [role="radiogroup"] label {
        min-height: 2.65rem;
        padding: 0.62rem 0.72rem;
        border-radius: 7px;
        transition: background 120ms ease;
    }
    [data-testid="stSidebar"] [role="radiogroup"] label:hover {background: #f3f7fd;}
    [data-testid="stSidebar"] [role="radiogroup"] label:has(input:checked) {
        background: var(--rfs-soft-blue);
        color: var(--rfs-blue-dark);
        font-weight: 650;
    }
    [data-testid="stSidebar"] [role="radiogroup"] label > div:first-child {display: none;}

    .rfs-brand {
        display: flex;
        align-items: center;
        gap: 0.7rem;
        padding: 0.2rem 0.25rem 1.2rem;
        font-weight: 750;
        color: var(--rfs-navy);
    }
    .rfs-logo {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        width: 2rem;
        height: 2rem;
        border-radius: 6px;
        background: var(--rfs-blue);
        color: white;
        font-size: 1rem;
        font-weight: 800;
    }
    .rfs-nav-label {
        margin: 0.6rem 0.25rem 0.3rem;
        color: #8a98aa;
        font-size: 0.7rem;
        font-weight: 750;
        text-transform: uppercase;
    }
    .rfs-help {
        margin-top: 1.2rem;
        padding: 0.9rem;
        border: 1px solid var(--rfs-border);
        border-radius: 8px;
        background: #f8fafc;
        color: var(--rfs-muted);
        font-size: 0.8rem;
        line-height: 1.45;
    }
    .rfs-topbar {
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 1rem;
        padding: 0.2rem 0 1rem;
        border-bottom: 1px solid var(--rfs-border);
        margin-bottom: 1.25rem;
    }
    .rfs-topbrand {
        display: flex;
        align-items: center;
        gap: 0.65rem;
        color: var(--rfs-navy) !important;
        font-weight: 750;
    }
    .rfs-status {
        padding: 0.35rem 0.65rem;
        border: 1px solid #b9dfd0;
        border-radius: 999px;
        background: #edf9f4;
        color: var(--rfs-green);
        font-size: 0.75rem;
        font-weight: 650;
    }
    .rfs-pagehead {margin-bottom: 1.2rem;}
    .rfs-pagehead h1 {margin: 0; font-size: 1.65rem !important;}
    .rfs-pagehead p {margin: 0.28rem 0 0; color: var(--rfs-muted);}

    div[data-testid="stVerticalBlockBorderWrapper"] {
        background: #ffffff;
        border-color: var(--rfs-border) !important;
        border-radius: 8px !important;
        box-shadow: 0 1px 3px rgba(15, 35, 65, 0.04);
    }
    [data-testid="stMetric"] {
        background: #ffffff;
        border: 1px solid var(--rfs-border);
        padding: 14px;
        border-radius: 8px;
        box-shadow: 0 1px 3px rgba(15, 35, 65, 0.04);
    }
    [data-testid="stFileUploaderDropzone"] {
        min-height: 9rem;
        background: #f8fbff;
        border: 1px dashed #a9bdd8;
        border-radius: 8px;
    }
    .stButton > button, .stDownloadButton > button, [data-testid="stFormSubmitButton"] button {
        min-height: 2.55rem;
        border-radius: 6px;
        font-weight: 650;
    }
    .stButton > button[kind="primary"], .stDownloadButton > button[kind="primary"],
    [data-testid="stFormSubmitButton"] button[kind="primary"] {
        background: var(--rfs-blue);
        border-color: var(--rfs-blue);
        color: #ffffff !important;
    }
    .stButton > button[kind="primary"] p, .stDownloadButton > button[kind="primary"] p,
    [data-testid="stFormSubmitButton"] button[kind="primary"] p {color: #ffffff !important;}
    .stButton > button[kind="primary"]:hover, .stDownloadButton > button[kind="primary"]:hover {
        background: var(--rfs-blue-dark);
        border-color: var(--rfs-blue-dark);
    }
    [data-testid="stDataFrame"] {
        border: 1px solid var(--rfs-border);
        border-radius: 8px;
        overflow: hidden;
    }
    .rfs-card-title {font-weight: 750; color: var(--rfs-navy); margin-bottom: 0.25rem;}
    .rfs-card-copy {min-height: 3.2rem; color: var(--rfs-muted); font-size: 0.86rem;}
    .rfs-kicker {
        display: inline-block;
        color: var(--rfs-blue-dark);
        background: var(--rfs-soft-blue);
        border-radius: 4px;
        padding: 0.22rem 0.42rem;
        font-size: 0.68rem;
        font-weight: 750;
        text-transform: uppercase;
        margin-bottom: 0.65rem;
    }
    .public-nav {
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 1rem;
        padding: 0.1rem 0 1rem;
        margin-bottom: 1rem;
        border-bottom: 1px solid var(--rfs-border);
    }
    .public-nav-links {display: flex; align-items: center; gap: 1.15rem;}
    .public-nav a {
        color: #42526a;
        font-size: 0.84rem;
        font-weight: 600;
        text-decoration: none;
    }
    .public-nav a:hover {color: var(--rfs-blue);}
    .public-nav .public-login {
        padding: 0.6rem 0.9rem;
        border-radius: 6px;
        background: var(--rfs-blue);
        color: #ffffff;
    }
    .public-hero {
        min-height: 520px;
        display: flex;
        align-items: center;
        padding: 3rem;
        border: 1px solid var(--rfs-border);
        border-radius: 8px;
        background-position: center;
        background-size: cover;
        overflow: hidden;
    }
    .public-hero-copy {width: min(47%, 540px);}
    .public-hero h1 {
        max-width: 520px;
        margin: 0 0 1rem;
        font-size: 3rem !important;
        line-height: 1.05;
    }
    .public-hero p {
        max-width: 510px;
        margin: 0 0 1.5rem;
        color: #34455d;
        font-size: 1rem;
        line-height: 1.65;
    }
    .public-actions {display: flex; flex-wrap: wrap; gap: 0.7rem;}
    .public-actions a {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        min-height: 2.75rem;
        padding: 0.65rem 1rem;
        border: 1px solid var(--rfs-border);
        border-radius: 6px;
        background: #ffffff;
        color: var(--rfs-navy);
        font-size: 0.86rem;
        font-weight: 700;
        text-decoration: none;
    }
    .public-actions a:first-child {
        border-color: var(--rfs-blue);
        background: var(--rfs-blue);
        color: #ffffff;
    }
    .public-section {padding: 3rem 0 1rem;}
    .public-section-head {max-width: 660px; margin-bottom: 1.4rem;}
    .public-section-head h2 {margin: 0 0 0.45rem; font-size: 1.65rem !important;}
    .public-section-head p {margin: 0; color: var(--rfs-muted);}
    .public-card {
        min-height: 180px;
        padding: 1.25rem;
        border: 1px solid var(--rfs-border);
        border-radius: 8px;
        background: #ffffff;
        box-shadow: 0 1px 3px rgba(15, 35, 65, 0.04);
    }
    .public-card-number {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        width: 2rem;
        height: 2rem;
        margin-bottom: 1rem;
        border-radius: 6px;
        background: var(--rfs-soft-blue);
        color: var(--rfs-blue-dark);
        font-size: 0.8rem;
        font-weight: 800;
    }
    .public-card h3 {margin: 0 0 0.45rem;}
    .public-card p {margin: 0; color: var(--rfs-muted); font-size: 0.88rem; line-height: 1.55;}
    .public-band {
        margin-top: 2.5rem;
        padding: 1.6rem 1.8rem;
        border: 1px solid #bcd0ea;
        border-radius: 8px;
        background: #edf5ff;
    }
    .public-band h2 {margin: 0 0 0.45rem;}
    .public-band p {max-width: 850px; margin: 0; color: #42526a;}
    .public-footer {
        display: flex;
        justify-content: space-between;
        gap: 1rem;
        margin-top: 3rem;
        padding: 1.4rem 0 0;
        border-top: 1px solid var(--rfs-border);
        color: var(--rfs-muted);
        font-size: 0.8rem;
    }
    .status-ok {padding: 12px 14px; border-left: 4px solid #17864b; background: #eff8f3;}
    .status-error {padding: 12px 14px; border-left: 4px solid #c73838; background: #fff2f2;}

    @media (max-width: 768px) {
        .block-container {padding: 4rem 0.9rem 3rem;}
        .rfs-topbar {align-items: flex-start;}
        .rfs-status {display: none;}
        .rfs-card-copy {min-height: 0;}
        .public-nav-links a:not(.public-login) {display: none;}
        .public-hero {min-height: 570px; align-items: flex-start; padding: 2rem 1.25rem; background-position: 66% center;}
        .public-hero-copy {width: 100%; padding: 0.9rem; border-radius: 7px; background: rgba(255,255,255,0.92);}
        .public-hero h1 {font-size: 2.15rem !important;}
        .public-footer {flex-direction: column;}
        [data-testid="column"] {min-width: 100% !important;}
    }
    </style>
    """,
    unsafe_allow_html=True,
)


def render_public_home() -> None:
    hero_path = APP_DIR / "assets" / "accounting-workspace-hero.png"
    hero_data = base64.b64encode(hero_path.read_bytes()).decode("ascii")

    st.markdown(
        '<nav class="public-nav">'
        '<div class="rfs-topbrand"><span class="rfs-logo">R</span>'
        "<span>Raman Financial Services</span></div>"
        '<div class="public-nav-links">'
        '<a href="?view=home">Home</a>'
        '<a href="#public-tools">Tools</a>'
        '<a href="#public-workflow">Workflow</a>'
        '<a class="public-login" href="?view=workspace">Open secure tools</a>'
        "</div></nav>",
        unsafe_allow_html=True,
    )
    st.markdown(
        f'<section class="public-hero" style="background-image:url(data:image/png;base64,{hero_data})">'
        '<div class="public-hero-copy">'
        '<span class="rfs-kicker">Accounting and financial tools</span>'
        "<h1>Financial work, organized.</h1>"
        "<p>Prepare bank transactions, automatic bookkeeping categories, annual workbooks, payroll records, compiled financial "
        "statements, mortgage estimates, and real estate investment analysis in one workspace.</p>"
        '<div class="public-actions">'
        '<a href="?view=workspace">Open secure tools</a>'
        '<a href="#public-tools">Explore tools</a>'
        "</div></div></section>",
        unsafe_allow_html=True,
    )

    st.markdown(
        '<section class="public-section" id="public-tools">'
        '<div class="public-section-head"><span class="rfs-kicker">Tools</span>'
        "<h2>One workspace for recurring financial work</h2>"
        "<p>Move from source documents to reviewable working files without changing systems.</p>"
        "</div></section>",
        unsafe_allow_html=True,
    )
    public_tools = [
        ("01", "Statement extraction", "Convert supported bank and credit-card statements into organized transactions."),
        ("02", "Auto categorization", "Apply reviewable bookkeeping categories using merchant and transaction-direction rules."),
        ("03", "Annual workbooks", "Combine verified monthly files while preserving statement order and separation."),
        ("04", "Payroll records", "Calculate payroll, update annual registers, and prepare employee payslips."),
        ("05", "Financial statements", "Create reviewable draft statements from Schedule 100 and Schedule 125 PDFs."),
        ("06", "Mortgage planning", "Estimate mortgage capacity using the Gross Debt Service calculation."),
        ("07", "Investment analysis", "Compare property costs, cash flow, rents, and supporting market inputs."),
    ]
    for row_start in range(0, len(public_tools), 3):
        columns = st.columns(3)
        for column, (number, title, copy) in zip(columns, public_tools[row_start : row_start + 3]):
            with column:
                st.markdown(
                    f'<article class="public-card"><span class="public-card-number">{number}</span>'
                    f"<h3>{title}</h3><p>{copy}</p></article>",
                    unsafe_allow_html=True,
                )

    st.markdown(
        '<section class="public-section" id="public-workflow">'
        '<div class="public-section-head"><span class="rfs-kicker">Workflow</span>'
        "<h2>Built around review, not blind automation</h2>"
        "<p>Each workflow keeps the original documents central to the accounting review.</p>"
        "</div></section>",
        unsafe_allow_html=True,
    )
    workflow_columns = st.columns(3)
    workflow_steps = [
        ("1", "Prepare", "Upload statements, schedules, workbooks, or property inputs."),
        ("2", "Review", "Reconcile totals and compare generated records with source documents."),
        ("3", "Export", "Download the completed Excel, PDF, or Word working file."),
    ]
    for column, (number, title, copy) in zip(workflow_columns, workflow_steps):
        with column:
            st.markdown(
                f'<article class="public-card"><span class="public-card-number">{number}</span>'
                f"<h3>{title}</h3><p>{copy}</p></article>",
                unsafe_allow_html=True,
            )

    st.markdown(
        '<section class="public-band"><h2>Public website. Protected financial workspace.</h2>'
        "<p>This home page is open to everyone. Bank statements, payroll, financial statements, "
        "mortgage calculations, and investment tools remain behind the workspace password.</p>"
        '<div class="public-actions" style="margin-top:1rem">'
        '<a href="?view=workspace">Sign in to the workspace</a></div></section>'
        '<footer class="public-footer"><strong>Raman Financial Services</strong>'
        "<span>Financial tools for organized, reviewable work.</span></footer>",
        unsafe_allow_html=True,
    )


def require_password() -> None:
    """Protect hosted deployments when APP_PASSWORD is configured."""
    expected = os.environ.get("APP_PASSWORD", "")
    if not expected:
        try:
            expected = str(st.secrets.get("APP_PASSWORD", ""))
        except FileNotFoundError:
            expected = ""
    if not expected or st.session_state.get("authenticated"):
        return

    st.markdown(
        '<div class="rfs-topbar"><div class="rfs-topbrand">'
        '<span class="rfs-logo">R</span><span>Raman Financial Services</span></div>'
        '<span class="rfs-status">Secure workspace</span></div>',
        unsafe_allow_html=True,
    )
    intro_col, login_col = st.columns([1.35, 0.85], gap="large")
    with intro_col:
        st.markdown('<span class="rfs-kicker">Accounting workspace</span>', unsafe_allow_html=True)
        st.title("Work with financial records in one secure place.")
        st.markdown(
            "Bank extraction, annual workbooks, payroll, financial statements, "
            "mortgage qualification, and real estate investment analysis."
        )
        benefit_cols = st.columns(3)
        benefit_cols[0].markdown("**Private**\n\nPassword-protected access")
        benefit_cols[1].markdown("**Practical**\n\nFiles ready for review")
        benefit_cols[2].markdown("**Connected**\n\nOne consistent workspace")
    with login_col:
        with st.container(border=True):
            st.subheader("Sign in")
            st.caption("Enter your workspace password.")
            password = st.text_input("Password", type="password")
            if st.button("Sign in", type="primary", use_container_width=True):
                if hmac.compare_digest(password, expected):
                    st.session_state["authenticated"] = True
                    st.rerun()
                else:
                    st.error("Incorrect password.")
            st.markdown("[Back to public home](?view=home)")
    st.stop()


requested_view = str(st.query_params.get("view", "home")).lower()
if requested_view != "workspace":
    render_public_home()
    st.stop()

require_password()


SUPPORTED_BANKS = ["Auto-detect", "BMO", "CIBC", "RBC", "Tangerine", "Vancity", "TD", "Other bank"]
PAY_PERIODS = {"Weekly": 52, "Biweekly": 26, "Semi-monthly": 24, "Monthly": 12}
FED_BRACKETS_2026 = [
    (58523, 0.14),
    (117045, 0.205),
    (181440, 0.26),
    (258482, 0.29),
    (float("inf"), 0.33),
]
BC_BRACKETS_2026 = [
    (50363, 0.0506),
    (100728, 0.0770),
    (115648, 0.1050),
    (140430, 0.1229),
    (190405, 0.1470),
    (265545, 0.1680),
    (float("inf"), 0.2050),
]
FED_BPA_2026 = 16452
BC_BPA_2026 = 13216
CPP1_RATE = 0.0595
CPP2_RATE = 0.04
CPP_BASIC_EXEMPTION = 3500
CPP1_YMPE = 71300
CPP2_YAMPE = 81200
EI_RATE = 0.0163
EI_MIE = 68900
PAYROLL_COLUMNS = [
    "employee_id",
    "pay_start",
    "pay_end",
    "pay_date",
    "hours",
    "rate",
    "salary_amount",
    "overtime_hours",
    "overtime_rate",
    "overtime_pay",
    "regular_pay",
    "stat_pay",
    "sick_pay",
    "vacation_pay",
    "vac_accrual",
    "gross",
    "cpp",
    "ei",
    "tax_fed",
    "tax_prov",
    "other_d",
    "reimb",
    "net",
    "ytd_gross",
    "ytd_cpp",
    "ytd_ei",
    "ytd_tax_fed",
    "ytd_tax_prov",
    "ytd_other_d",
    "ytd_reimb",
    "ytd_net",
    "ytd_regular_pay",
    "ytd_stat_pay",
    "ytd_sick_pay",
    "ytd_vacation_pay",
    "ytd_overtime_pay",
    "ytd_vac_accrual",
    "ytd_vac_paid",
    "pdf_link",
    "status",
]
DEFAULT_REALTOR_URL = (
    "https://www.realtor.ca/map#ZoomLevel=15&Center=49.158114%2C-122.846239"
    "&LatitudeMax=49.16707&LongitudeMax=-122.81311&LatitudeMin=49.14916"
    "&LongitudeMin=-122.87937&Sort=6-D&PropertyTypeGroupID=1"
    "&TransactionTypeId=2&PropertySearchTypeId=1&OwnershipTypeGroupId=1&Currency=CAD"
)


def safe_name(value: str) -> str:
    return "".join(char if char.isalnum() or char in "-_. " else "_" for char in value).strip()


def docling_json_for_pdf(pdf_path: Path, work_dir: Path) -> Path:
    try:
        from docling.datamodel.base_models import InputFormat
        from docling.datamodel.pipeline_options import (
            PdfPipelineOptions,
            TesseractCliOcrOptions,
        )
        from docling.document_converter import DocumentConverter, PdfFormatOption
    except ImportError as exc:
        raise RuntimeError(
            "BMO PDF conversion requires Docling. Install the website requirements and restart the app."
        ) from exc

    json_path = work_dir / f"{pdf_path.stem}.docling.json"
    pipeline_options = PdfPipelineOptions()
    pipeline_options.do_ocr = True
    pipeline_options.do_table_structure = True
    pipeline_options.ocr_options = TesseractCliOcrOptions(
        lang=["eng"],
        force_full_page_ocr=True,
        psm=6,
    )
    converter = DocumentConverter(
        format_options={
            InputFormat.PDF: PdfFormatOption(pipeline_options=pipeline_options),
        }
    )
    result = converter.convert(str(pdf_path))
    document = result.document
    if hasattr(document, "export_to_dict"):
        data = document.export_to_dict()
    elif hasattr(document, "model_dump"):
        data = document.model_dump()
    else:
        raise RuntimeError("This Docling version could not export JSON.")
    json_path.write_text(json.dumps(data, indent=2), encoding="utf-8")
    return json_path


def reconciliation_status(df: pd.DataFrame) -> dict:
    """Check whether extracted transaction arithmetic reaches the closing balance."""
    opening = df.loc[df["Category"] == "Opening Balance", "Balance"].dropna()
    closing = df.loc[df["Category"] == "Closing Totals", "Balance"].dropna()
    normal = df[~df["Category"].isin(["Opening Balance", "Closing Totals"])]

    if opening.empty or closing.empty:
        return {
            "status": "review",
            "message": "Opening or closing balance was not found. Compare the Excel totals with the statement.",
        }

    total_debits = float(pd.to_numeric(normal["Debit"], errors="coerce").fillna(0).sum())
    total_credits = float(pd.to_numeric(normal["Credit"], errors="coerce").fillna(0).sum())
    credit_card_statement = df["Description"].astype(str).str.contains(
        "Previous Statement Balance", case=False, regex=False
    ).any()
    if credit_card_statement:
        calculated = round(float(opening.iloc[0]) + total_debits - total_credits, 2)
    else:
        calculated = round(float(opening.iloc[0]) - total_debits + total_credits, 2)
    expected = round(float(closing.iloc[-1]), 2)
    if abs(calculated - expected) <= 0.02:
        return {
            "status": "reconciled",
            "message": f"Reconciled to the statement closing balance (${expected:,.2f}).",
        }
    return {
        "status": "review",
        "message": (
            f"Needs review: calculated closing balance is ${calculated:,.2f}, "
            f"but the statement shows ${expected:,.2f}."
        ),
    }


def process_statement(uploaded_file, selected_bank: str) -> dict:
    work_dir = Path(tempfile.mkdtemp(prefix="bank-extractor-"))
    input_path = work_dir / safe_name(uploaded_file.name)
    input_path.write_bytes(uploaded_file.getvalue())

    bank = extractor.detect_bank_from_file(input_path) if selected_bank == "Auto-detect" else selected_bank
    if bank == "Unknown":
        bank = "Other bank"

    rows = []
    direct_error = None
    scanned_ocr_attempted = False
    image_only_pdf = (
        input_path.suffix.lower() == ".pdf"
        and not extractor.read_pdf_text(input_path).strip()
    )
    if image_only_pdf:
        scanned_ocr_attempted = True
        try:
            rows = extractor.extract_scanned_debit_credit_pdf_transactions(input_path)
        except RuntimeError as exc:
            direct_error = exc

    if input_path.suffix.lower() == ".pdf" and not image_only_pdf:
        try:
            rows = extractor.extract_from_statement_file(input_path, bank)
        except RuntimeError as exc:
            direct_error = exc

    if not rows:
        if scanned_ocr_attempted:
            if direct_error is not None:
                raise RuntimeError(
                    f"The scanned statement reader could not process this file: {direct_error}"
                ) from direct_error
            raise RuntimeError(
                "The scanned statement reader did not find recognizable Date, Description, "
                "Debit/Withdrawal, Credit/Deposit, and Balance columns. No Excel file was created."
            )
        effective_path = input_path
        if input_path.suffix.lower() == ".pdf":
            detail = f": {direct_error}" if direct_error is not None else ""
            raise RuntimeError(
                "This PDF layout did not match a safe local extractor. The website did not "
                f"start the memory-heavy Docling fallback{detail}"
            )
        try:
            rows = extractor.extract_from_statement_file(effective_path, bank)
        except Exception as exc:
            if direct_error is not None:
                raise RuntimeError(f"Direct extraction failed: {direct_error}. Docling fallback failed: {exc}") from exc
            raise
    if not rows:
        raise RuntimeError(f"No transactions were found for {bank}.")

    df = extractor.to_dataframe(rows)
    output_name = f"{input_path.stem}_{bank}_transactions.xlsx"
    output_path = extractor.write_excel(df, work_dir / output_name)
    summary = extractor.build_summary(df)
    normal_df = df[~df["Category"].isin(["Opening Balance", "Closing Totals"])]
    visible_transactions = extractor.to_signed_amount_view(normal_df)

    metrics = {row["Metric"]: row["Value"] for _, row in summary.iterrows()}
    reconciliation = reconciliation_status(df)
    is_rbc_chequing = bank == "RBC" and not df["Description"].astype(str).str.contains(
        "Previous Statement Balance", case=False, regex=False
    ).any()
    if is_rbc_chequing and reconciliation["status"] != "reconciled":
        raise RuntimeError(
            "RBC safety check failed. The extracted transactions do not match the "
            "statement totals and closing balance, so no Excel file was produced. "
            "This statement needs review before it can be exported."
        )
    return {
        "bank": bank,
        "source": uploaded_file.name,
        "output_name": output_path.name,
        "output_bytes": output_path.read_bytes(),
        "transactions": visible_transactions,
        "summary": summary,
        "metrics": metrics,
        "reconciliation": reconciliation,
    }


def make_zip(results: list[dict]) -> bytes:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        for result in results:
            archive.writestr(result["output_name"], result["output_bytes"])
    return buffer.getvalue()


def merge_uploaded_excels(files, output_name: str) -> tuple[bytes, pd.DataFrame, pd.DataFrame, str]:
    work_dir = Path(tempfile.mkdtemp(prefix="annual-merge-"))
    paths = []
    for file in files:
        path = work_dir / safe_name(file.name)
        path.write_bytes(file.getvalue())
        paths.append(path)

    output_name = safe_name(output_name or "Annual_transactions.xlsx")
    if not output_name.lower().endswith(".xlsx"):
        output_name += ".xlsx"
    output_path = extractor.merge_excel_files(paths, work_dir / output_name)
    transactions = pd.read_excel(output_path, sheet_name="Annual Transactions")
    summary = pd.read_excel(output_path, sheet_name="Annual Summary")
    return output_path.read_bytes(), transactions, summary, output_path.name


def merge_extracted_results(results: list[dict], output_name: str) -> tuple[bytes, pd.DataFrame, pd.DataFrame, str]:
    work_dir = Path(tempfile.mkdtemp(prefix="annual-from-extract-"))
    paths = []
    for index, result in enumerate(results, start=1):
        path = work_dir / f"{index:02d}_{safe_name(result['output_name'])}"
        path.write_bytes(result["output_bytes"])
        paths.append(path)

    output_name = safe_name(output_name or "Annual_transactions.xlsx")
    if not output_name.lower().endswith(".xlsx"):
        output_name += ".xlsx"
    output_path = extractor.merge_excel_files(paths, work_dir / output_name)
    transactions = pd.read_excel(output_path, sheet_name="Annual Transactions")
    summary = pd.read_excel(output_path, sheet_name="Annual Summary")
    return output_path.read_bytes(), transactions, summary, output_path.name


def progressive_tax(annual_income: float, brackets: list[tuple[float, float]]) -> float:
    tax = 0.0
    previous_limit = 0.0
    for limit, rate in brackets:
        taxable = min(annual_income, limit) - previous_limit
        if taxable > 0:
            tax += taxable * rate
        if annual_income <= limit:
            break
        previous_limit = limit
    return max(0.0, tax)


def calculate_payroll(values: dict) -> dict:
    periods = PAY_PERIODS[values["frequency"]]
    regular_pay = values["hours"] * values["rate"]
    overtime_pay = values["overtime_hours"] * values["overtime_rate"]
    gross = (
        regular_pay
        + values["salary_amount"]
        + overtime_pay
        + values["stat_pay"]
        + values["sick_pay"]
        + values["vacation_pay"]
        + values["bonus"]
    )

    annualized_gross = gross * periods
    cpp1_base = max(0.0, min(annualized_gross, CPP1_YMPE) - CPP_BASIC_EXEMPTION)
    cpp1 = min(cpp1_base * CPP1_RATE / periods, max(0.0, (CPP1_YMPE - CPP_BASIC_EXEMPTION) * CPP1_RATE - values["ytd_cpp"]))
    cpp2_base = max(0.0, min(annualized_gross, CPP2_YAMPE) - CPP1_YMPE)
    cpp2 = min(cpp2_base * CPP2_RATE / periods, max(0.0, (CPP2_YAMPE - CPP1_YMPE) * CPP2_RATE - values["ytd_cpp2"]))
    cpp = round(max(0.0, cpp1 + cpp2), 2)

    ei_max = EI_MIE * EI_RATE
    ei = round(min(gross * EI_RATE, max(0.0, ei_max - values["ytd_ei"])), 2)

    federal_annual = progressive_tax(annualized_gross, FED_BRACKETS_2026)
    federal_credit = FED_BPA_2026 * 0.14
    tax_fed = round(max(0.0, federal_annual - federal_credit) / periods, 2)

    bc_annual = progressive_tax(annualized_gross, BC_BRACKETS_2026)
    bc_credit = BC_BPA_2026 * 0.0506
    tax_prov = round(max(0.0, bc_annual - bc_credit) / periods, 2)

    deductions = cpp + ei + tax_fed + tax_prov + values["other_deductions"]
    net = round(gross - deductions + values["reimbursements"], 2)
    return {
        "regular_pay": round(regular_pay, 2),
        "overtime_pay": round(overtime_pay, 2),
        "gross": round(gross, 2),
        "cpp": cpp,
        "ei": ei,
        "tax_fed": tax_fed,
        "tax_prov": tax_prov,
        "other_deductions": round(values["other_deductions"], 2),
        "reimbursements": round(values["reimbursements"], 2),
        "total_deductions": round(deductions, 2),
        "net": net,
        "employer_cpp": cpp,
        "employer_ei": round(ei * 1.4, 2),
    }


def _pdoc_amount(text: str, label: str) -> float | None:
    pattern = rf"{re.escape(label)}\s+(-?\d[\d,]*\.\d{{2}})"
    match = re.search(pattern, text, flags=re.IGNORECASE)
    if not match:
        return None
    return float(match.group(1).replace(",", ""))


def extract_pdoc_result(uploaded_file) -> dict:
    if uploaded_file is None:
        return {}
    try:
        from pypdf import PdfReader
    except ImportError as exc:
        raise RuntimeError("CRA PDOC PDF import requires pypdf. Install requirements and restart the app.") from exc

    try:
        reader = PdfReader(io.BytesIO(uploaded_file.getvalue()))
        text = "\n".join(page.extract_text() or "" for page in reader.pages)
    except Exception as exc:
        raise RuntimeError(f"Could not read CRA PDOC PDF: {exc}") from exc

    if "Payroll Deductions Online Calculator" not in text:
        raise RuntimeError("This does not look like a CRA PDOC result PDF.")

    values = {
        "gross": _pdoc_amount(text, "Salary or wages income"),
        "tax_fed": _pdoc_amount(text, "Federal tax deduction"),
        "tax_prov": _pdoc_amount(text, "Provincial tax deduction"),
        "cpp": _pdoc_amount(text, "CPP deductions"),
        "ei": _pdoc_amount(text, "EI deductions"),
        "total_deductions": _pdoc_amount(text, "Total deductions"),
        "net": _pdoc_amount(text, "Net amount"),
    }
    missing = [name for name, amount in values.items() if amount is None and name != "gross"]
    if missing:
        raise RuntimeError("Could not find all CRA PDOC deduction amounts in this PDF.")

    date_match = re.search(r"Date the employee is paid:\s+(\d{4}-\d{2}-\d{2})", text)
    frequency_match = re.search(r"Pay period frequency:\s+([^\n(]+)", text)
    values["pay_date"] = date_match.group(1) if date_match else ""
    values["frequency"] = frequency_match.group(1).strip() if frequency_match else ""
    return values


def apply_pdoc_to_calc(calc: dict, pdoc: dict, values: dict) -> dict:
    updated = dict(calc)
    for field in ["cpp", "ei", "tax_fed", "tax_prov", "total_deductions", "net"]:
        if pdoc.get(field) is not None:
            updated[field] = round(float(pdoc[field]), 2)
    updated["employer_cpp"] = updated["cpp"]
    updated["employer_ei"] = round(updated["ei"] * 1.4, 2)
    updated["pdoc_source"] = "CRA PDOC PDF"
    if pdoc.get("gross") is not None and abs(float(pdoc["gross"]) - calc["gross"]) > 0.01:
        updated["pdoc_warning"] = (
            f"PDOC gross pay is ${float(pdoc['gross']):,.2f}, but the payroll form gross pay is ${calc['gross']:,.2f}."
        )
    elif pdoc.get("frequency") and values.get("frequency") and pdoc["frequency"].lower() != values["frequency"].lower():
        updated["pdoc_warning"] = f"PDOC frequency is {pdoc['frequency']}, but the form frequency is {values['frequency']}."
    else:
        updated["pdoc_warning"] = ""
    return updated


def load_payroll_register(uploaded_file) -> pd.DataFrame:
    if uploaded_file is None:
        return pd.DataFrame(columns=PAYROLL_COLUMNS)
    try:
        workbook = pd.ExcelFile(uploaded_file)
        sheet_name = "Payroll" if "Payroll" in workbook.sheet_names else workbook.sheet_names[0]
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
    except Exception as exc:
        raise RuntimeError(f"Could not read payroll register: {exc}") from exc
    for column in PAYROLL_COLUMNS:
        if column not in df.columns:
            df[column] = pd.NA
    return df[PAYROLL_COLUMNS]


def payroll_ytd_before(register: pd.DataFrame, employee_id: str) -> dict:
    if register.empty or not employee_id:
        base = register
    else:
        base = register[register["employee_id"].astype(str).str.strip() == str(employee_id).strip()]
    fields = [
        "gross",
        "cpp",
        "ei",
        "tax_fed",
        "tax_prov",
        "other_d",
        "reimb",
        "net",
        "regular_pay",
        "stat_pay",
        "sick_pay",
        "vacation_pay",
        "overtime_pay",
        "vac_accrual",
    ]
    return {
        field: float(pd.to_numeric(base.get(field, pd.Series(dtype=float)), errors="coerce").fillna(0).sum())
        for field in fields
    }


def make_payroll_register_row(values: dict, calc: dict, ytd_before: dict) -> dict:
    ytd = {
        "gross": ytd_before["gross"] + calc["gross"],
        "cpp": ytd_before["cpp"] + calc["cpp"],
        "ei": ytd_before["ei"] + calc["ei"],
        "tax_fed": ytd_before["tax_fed"] + calc["tax_fed"],
        "tax_prov": ytd_before["tax_prov"] + calc["tax_prov"],
        "other_d": ytd_before["other_d"] + calc["other_deductions"],
        "reimb": ytd_before["reimb"] + calc["reimbursements"],
        "net": ytd_before["net"] + calc["net"],
        "regular_pay": ytd_before["regular_pay"] + calc["regular_pay"],
        "stat_pay": ytd_before["stat_pay"] + values["stat_pay"],
        "sick_pay": ytd_before["sick_pay"] + values["sick_pay"],
        "vacation_pay": ytd_before["vacation_pay"] + values["vacation_pay"],
        "overtime_pay": ytd_before["overtime_pay"] + calc["overtime_pay"],
        "vac_accrual": ytd_before["vac_accrual"] + values["vac_accrual"],
    }
    row = {column: pd.NA for column in PAYROLL_COLUMNS}
    row.update(
        {
            "employee_id": values["employee_id"],
            "pay_start": values["pay_start"],
            "pay_end": values["pay_end"],
            "pay_date": values["pay_date"],
            "hours": values["hours"],
            "rate": values["rate"],
            "salary_amount": values["salary_amount"],
            "overtime_hours": values["overtime_hours"],
            "overtime_rate": values["overtime_rate"],
            "overtime_pay": calc["overtime_pay"],
            "regular_pay": calc["regular_pay"],
            "stat_pay": values["stat_pay"],
            "sick_pay": values["sick_pay"],
            "vacation_pay": values["vacation_pay"],
            "vac_accrual": values["vac_accrual"],
            "gross": calc["gross"],
            "cpp": calc["cpp"],
            "ei": calc["ei"],
            "tax_fed": calc["tax_fed"],
            "tax_prov": calc["tax_prov"],
            "other_d": calc["other_deductions"],
            "reimb": calc["reimbursements"],
            "net": calc["net"],
            "ytd_gross": round(ytd["gross"], 2),
            "ytd_cpp": round(ytd["cpp"], 2),
            "ytd_ei": round(ytd["ei"], 2),
            "ytd_tax_fed": round(ytd["tax_fed"], 2),
            "ytd_tax_prov": round(ytd["tax_prov"], 2),
            "ytd_other_d": round(ytd["other_d"], 2),
            "ytd_reimb": round(ytd["reimb"], 2),
            "ytd_net": round(ytd["net"], 2),
            "ytd_regular_pay": round(ytd["regular_pay"], 2),
            "ytd_stat_pay": round(ytd["stat_pay"], 2),
            "ytd_sick_pay": round(ytd["sick_pay"], 2),
            "ytd_vacation_pay": round(ytd["vacation_pay"], 2),
            "ytd_overtime_pay": round(ytd["overtime_pay"], 2),
            "ytd_vac_accrual": round(ytd["vac_accrual"], 2),
            "ytd_vac_paid": round(ytd["vacation_pay"], 2),
            "status": "Calculated",
        }
    )
    return row


def export_payroll_register(register: pd.DataFrame) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        register.to_excel(writer, sheet_name="Payroll", index=False)
        summary = (
            register.groupby("employee_id", dropna=False)[["gross", "cpp", "ei", "tax_fed", "tax_prov", "net"]]
            .sum(numeric_only=True)
            .reset_index()
        )
        summary.to_excel(writer, sheet_name="Payroll Summary", index=False)
        for sheet_name, df in {"Payroll": register, "Payroll Summary": summary}.items():
            extractor.auto_adjust_columns(writer, sheet_name, df)
    return buffer.getvalue()


def build_payslip_pdf(company: dict, employee: dict, payroll: dict, calc: dict) -> bytes:
    try:
        from reportlab.lib import colors
        from reportlab.lib.pagesizes import letter
        from reportlab.lib.styles import getSampleStyleSheet
        from reportlab.lib.units import inch
        from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
    except ImportError as exc:
        raise RuntimeError("Payslip PDF requires reportlab. Install requirements and restart the app.") from exc

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=0.6 * inch, leftMargin=0.6 * inch)
    styles = getSampleStyleSheet()
    story = [
        Paragraph(company["name"], styles["Title"]),
        Paragraph(company["address"], styles["Normal"]),
        Spacer(1, 12),
        Paragraph("Payslip", styles["Heading2"]),
    ]

    details = [
        ["Employee", employee["name"], "Pay date", str(payroll["pay_date"])],
        ["Position", employee["position"], "Pay period", f"{payroll['pay_start']} to {payroll['pay_end']}"],
        ["Frequency", payroll["frequency"], "Province", "BC"],
    ]
    earnings = [
        ["Earnings", "Amount"],
        ["Regular pay", calc["regular_pay"]],
        ["Overtime pay", calc["overtime_pay"]],
        ["Stat pay", payroll["stat_pay"]],
        ["Vacation pay", payroll["vacation_pay"]],
        ["Bonus", payroll["bonus"]],
        ["Gross pay", calc["gross"]],
    ]
    deductions = [
        ["Deductions", "Amount"],
        ["CPP", payroll["cpp"]],
        ["EI", payroll["ei"]],
        ["Federal tax", payroll["tax_fed"]],
        ["Provincial tax", payroll["tax_prov"]],
        ["Other deductions", payroll["other_deductions"]],
        ["Total deductions", payroll["total_deductions"]],
        ["Reimbursements", payroll["reimbursements"]],
        ["Net pay", payroll["net"]],
    ]

    def money_table(rows: list[list]) -> Table:
        formatted = [[row[0], row[1] if isinstance(row[1], str) else f"${row[1]:,.2f}"] for row in rows]
        table = Table(formatted, colWidths=[3.0 * inch, 2.0 * inch])
        table.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E8EEF7")),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#CAD2E0")),
                    ("ALIGN", (1, 1), (1, -1), "RIGHT"),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                    ("TOPPADDING", (0, 0), (-1, -1), 6),
                ]
            )
        )
        return table

    details_table = Table(details, colWidths=[1.2 * inch, 2.3 * inch, 1.2 * inch, 2.0 * inch])
    details_table.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#CAD2E0"))]))
    story.extend([details_table, Spacer(1, 14), money_table(earnings), Spacer(1, 14), money_table(deductions)])
    story.append(Spacer(1, 12))
    story.append(Paragraph("Review payroll deductions against CRA PDOC before filing remittances.", styles["Italic"]))
    doc.build(story)
    return buffer.getvalue()


def build_payslip_excel(company: dict, employee: dict, payroll: dict, calc: dict) -> bytes:
    """Create an editable payroll statement matching the PDF payslip."""
    rows = [
        ["Company", company["name"], "", ""],
        ["Company address", company["address"], "", ""],
        ["Employee", employee["name"], "Pay date", payroll["pay_date"]],
        ["Position", employee["position"], "Pay period", f"{payroll['pay_start']} to {payroll['pay_end']}"],
        ["Frequency", payroll["frequency"], "Province", employee.get("province", "British Columbia")],
        ["", "", "", ""],
        ["Earnings", "Amount", "Deductions", "Amount"],
        ["Regular pay", calc["regular_pay"], "CPP", calc["cpp"]],
        ["Overtime pay", calc["overtime_pay"], "EI", calc["ei"]],
        ["Stat pay", payroll["stat_pay"], "Federal tax", calc["tax_fed"]],
        ["Sick pay", payroll["sick_pay"], "Provincial tax", calc["tax_prov"]],
        ["Vacation pay", payroll["vacation_pay"], "Other deductions", calc["other_deductions"]],
        ["Bonus", payroll["bonus"], "Total deductions", calc["total_deductions"]],
        ["Gross pay", calc["gross"], "Reimbursements", calc["reimbursements"]],
        ["", "", "Net pay", calc["net"]],
        ["", "", "", ""],
        ["Employer CPP", calc["employer_cpp"], "Employer EI", calc["employer_ei"]],
        ["Review payroll deductions against CRA PDOC before remitting or filing.", "", "", ""],
    ]
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        pd.DataFrame(rows).to_excel(writer, sheet_name="Payslip", index=False, header=False)
        ws = writer.book["Payslip"]
        from openpyxl.styles import Alignment, Font, PatternFill

        fill = PatternFill("solid", fgColor="E8EEF7")
        for cell in ws[7]:
            cell.font = Font(bold=True)
            cell.fill = fill
        for row_number in range(8, 18):
            ws.cell(row=row_number, column=2).number_format = '$#,##0.00'
            ws.cell(row=row_number, column=4).number_format = '$#,##0.00'
        for column, width in {"A": 27, "B": 18, "C": 25, "D": 18}.items():
            ws.column_dimensions[column].width = width
        for row in ws.iter_rows():
            for cell in row:
                cell.alignment = Alignment(vertical="top")
        ws.freeze_panes = "A7"
    return buffer.getvalue()


def _pd7a_values(calc: dict) -> dict:
    income_tax = round(calc["tax_fed"] + calc["tax_prov"], 2)
    employee_cpp = round(calc["cpp"], 2)
    employer_cpp = round(calc["employer_cpp"], 2)
    employee_ei = round(calc["ei"], 2)
    employer_ei = round(calc["employer_ei"], 2)
    total_cpp = round(employee_cpp + employer_cpp, 2)
    total_ei = round(employee_ei + employer_ei, 2)
    return {
        "income_tax": income_tax,
        "employee_cpp": employee_cpp,
        "employer_cpp": employer_cpp,
        "employee_ei": employee_ei,
        "employer_ei": employer_ei,
        "total_cpp": total_cpp,
        "total_ei": total_ei,
        "total_remittance": round(income_tax + total_cpp + total_ei, 2),
    }


def build_pd7a_excel(company: dict, payroll: dict, calc: dict) -> bytes:
    """Create a PD7A-style remittance worksheet for CRA entry and review."""
    values = _pd7a_values(calc)
    rows = [
        ["PD7A-style payroll remittance summary", "", ""],
        ["Use this worksheet to complete CRA remittance entry or an official voucher.", "", ""],
        ["Employer name", company["name"], ""],
        ["Payroll account number", company.get("payroll_account", ""), ""],
        ["Remittance period", f"{payroll['pay_start']} to {payroll['pay_end']}", ""],
        ["Payment date", payroll["pay_date"], ""],
        ["", "", ""],
        ["Line", "Description", "Amount"],
        ["Gross payroll", "Total gross remuneration for the period", calc["gross"]],
        ["Income tax", "Federal and provincial income tax deducted", values["income_tax"]],
        ["Employee CPP", "CPP deducted from employee", values["employee_cpp"]],
        ["Employer CPP", "Employer CPP contribution", values["employer_cpp"]],
        ["Total CPP", "Employee CPP plus employer CPP", values["total_cpp"]],
        ["Employee EI", "EI deducted from employee", values["employee_ei"]],
        ["Employer EI", "Employer EI contribution", values["employer_ei"]],
        ["Total EI", "Employee EI plus employer EI", values["total_ei"]],
        ["Total remittance", "Income tax plus total CPP plus total EI", values["total_remittance"]],
        ["", "", ""],
        ["Review against CRA PDOC and payroll remittance records before paying.", "", ""],
    ]
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        pd.DataFrame(rows).to_excel(writer, sheet_name="PD7A Summary", index=False, header=False)
        ws = writer.book["PD7A Summary"]
        from openpyxl.styles import Alignment, Font, PatternFill

        header_fill = PatternFill("solid", fgColor="E8EEF7")
        total_fill = PatternFill("solid", fgColor="DDEBFF")
        ws["A1"].font = Font(bold=True, size=14)
        for cell in ws[8]:
            cell.font = Font(bold=True)
            cell.fill = header_fill
        for row_number in range(9, 18):
            ws.cell(row=row_number, column=3).number_format = '$#,##0.00'
        for cell in ws[17]:
            cell.font = Font(bold=True)
            cell.fill = total_fill
        for column, width in {"A": 24, "B": 54, "C": 18}.items():
            ws.column_dimensions[column].width = width
        for row in ws.iter_rows():
            for cell in row:
                cell.alignment = Alignment(vertical="top")
    return buffer.getvalue()


def build_pd7a_pdf(company: dict, payroll: dict, calc: dict) -> bytes:
    """Create a polished PD7A-style PDF summary for payroll records."""
    try:
        from reportlab.lib import colors
        from reportlab.lib.pagesizes import letter
        from reportlab.lib.styles import getSampleStyleSheet
        from reportlab.lib.units import inch
        from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
    except ImportError as exc:
        raise RuntimeError("PD7A PDF requires reportlab. Install requirements and restart the app.") from exc

    values = _pd7a_values(calc)
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        leftMargin=0.65 * inch,
        rightMargin=0.65 * inch,
        topMargin=0.65 * inch,
        bottomMargin=0.65 * inch,
    )
    styles = getSampleStyleSheet()
    navy = colors.HexColor("#16324F")
    teal = colors.HexColor("#087F8C")
    line = colors.HexColor("#D5DEE8")

    header = Table(
        [
            [
                Paragraph(
                    f"<font color='#FFFFFF' size='16'><b>{company['name']}</b></font><br/>"
                    f"<font color='#DDE8F1' size='8'>{company['address']}<br/>"
                    f"Payroll account: {company.get('payroll_account') or 'Not entered'}</font>",
                    styles["Normal"],
                ),
                Paragraph(
                    "<para alignment='right'><font color='#FFFFFF' size='18'><b>PD7A SUMMARY</b></font></para>",
                    styles["Normal"],
                ),
            ]
        ],
        colWidths=[4.4 * inch, 2.8 * inch],
    )
    header.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (0, 0), navy),
                ("BACKGROUND", (1, 0), (1, 0), teal),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 14),
                ("RIGHTPADDING", (0, 0), (-1, -1), 14),
                ("TOPPADDING", (0, 0), (-1, -1), 14),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 14),
            ]
        )
    )
    details = Table(
        [
            ["REMITTANCE PERIOD", f"{payroll['pay_start']} to {payroll['pay_end']}"],
            ["PAY DATE", str(payroll["pay_date"])],
            ["GROSS PAYROLL", f"${calc['gross']:,.2f}"],
        ],
        colWidths=[1.75 * inch, 5.45 * inch],
    )
    details.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#EEF3F8")),
                ("GRID", (0, 0), (-1, -1), 0.35, line),
                ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
                ("LEFTPADDING", (0, 0), (-1, -1), 10),
                ("RIGHTPADDING", (0, 0), (-1, -1), 10),
                ("TOPPADDING", (0, 0), (-1, -1), 8),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
            ]
        )
    )
    remittance = Table(
        [
            ["SOURCE DEDUCTION", "EMPLOYEE", "EMPLOYER", "TOTAL"],
            ["Income tax", f"${values['income_tax']:,.2f}", "-", f"${values['income_tax']:,.2f}"],
            ["CPP contributions", f"${values['employee_cpp']:,.2f}", f"${values['employer_cpp']:,.2f}", f"${values['total_cpp']:,.2f}"],
            ["EI premiums", f"${values['employee_ei']:,.2f}", f"${values['employer_ei']:,.2f}", f"${values['total_ei']:,.2f}"],
            ["TOTAL REMITTANCE DUE", "", "", f"${values['total_remittance']:,.2f}"],
        ],
        colWidths=[3.0 * inch, 1.4 * inch, 1.4 * inch, 1.4 * inch],
    )
    remittance.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), navy),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
                ("GRID", (0, 0), (-1, -1), 0.35, line),
                ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#EAF6F7")),
                ("TEXTCOLOR", (0, -1), (-1, -1), teal),
                ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
                ("LEFTPADDING", (0, 0), (-1, -1), 10),
                ("RIGHTPADDING", (0, 0), (-1, -1), 10),
                ("TOPPADDING", (0, 0), (-1, -1), 9),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 9),
            ]
        )
    )
    note = Paragraph(
        "<font color='#64748B' size='8'>PD7A-style summary for recordkeeping and CRA remittance entry. "
        "Review the amounts against CRA payroll records before making payment.</font>",
        styles["Normal"],
    )
    doc.build([header, Spacer(1, 12), details, Spacer(1, 14), remittance, Spacer(1, 12), note])
    return buffer.getvalue()


def save_uploaded_file(uploaded_file, folder: Path) -> Path | None:
    if uploaded_file is None:
        return None
    path = folder / safe_name(uploaded_file.name)
    path.write_bytes(uploaded_file.getvalue())
    return path


def run_real_estate_search(
    source_mode: str,
    realtor_url: str,
    listing_csv,
    zealty_upload,
    rental_upload,
    signal_upload,
    assumptions: InvestmentAssumptions,
) -> tuple[list[dict], dict]:
    with tempfile.TemporaryDirectory(prefix="real_estate_") as temp_name:
        temp_dir = Path(temp_name)
        listing_path = save_uploaded_file(listing_csv, temp_dir)
        zealty_path = save_uploaded_file(zealty_upload, temp_dir)
        rental_path = save_uploaded_file(rental_upload, temp_dir)
        signal_path = save_uploaded_file(signal_upload, temp_dir)

        if source_mode == "Saved listing CSV":
            if listing_path is None:
                raise ValueError("Upload a Realtor listing CSV.")
            realtor = RealtorSavedSearchProvider(csv_path=listing_path)
        else:
            if not realtor_url.strip():
                raise ValueError("Paste a Realtor.ca map search URL.")
            realtor = RealtorSavedSearchProvider(search_url=realtor_url.strip())

        zealty = ZealtyProvider(zealty_path)
        rentals = RentalProvider(rental_path)
        signals = SignalProvider(signal_path)
        rows = [
            score_property(
                listing,
                zealty.fetch(listing),
                rentals.fetch(listing),
                signals.fetch(listing),
                assumptions,
            )
            for listing in realtor.fetch()
        ]
        return merge_with_state(rows, {})


def real_estate_excel_bytes(database: pd.DataFrame, top: pd.DataFrame) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for sheet_name, frame in {
            "Property Database": database,
            "Top Opportunities": top,
        }.items():
            frame.to_excel(writer, sheet_name=sheet_name, index=False)
            worksheet = writer.sheets[sheet_name]
            worksheet.freeze_panes = "A2"
            worksheet.auto_filter.ref = worksheet.dimensions
            for column_cells in worksheet.columns:
                width = max(len(str(cell.value or "")) for cell in column_cells)
                worksheet.column_dimensions[column_cells[0].column_letter].width = min(max(width + 2, 11), 42)
    return buffer.getvalue()


NAV_ITEMS = [
    "Dashboard",
    "Extract statements",
    "Auto categorize",
    "Build annual file",
    "Payroll template",
    "Financial statements",
    "Maximum mortgage",
    "Real estate agent",
    "Guide",
]
PAGE_COPY = {
    "Dashboard": "Choose a workspace and continue your accounting work.",
    "Extract statements": "Upload bank statements and review extracted transactions.",
    "Auto categorize": "Apply reviewable bookkeeping categories to Excel or CSV transactions.",
    "Build annual file": "Combine verified monthly workbooks in statement order.",
    "Payroll template": "Calculate payroll and maintain the annual payroll register.",
    "Financial statements": "Prepare draft compiled financial statements from T2 schedules.",
    "Maximum mortgage": "Estimate mortgage capacity under the Gross Debt Service ratio.",
    "Real estate agent": "Compare listings, financing costs, cash flow, and investment signals.",
    "Guide": "File naming, review steps, and operating notes.",
}


def go_to_page(page: str) -> None:
    st.session_state["main_nav"] = page


def sign_out() -> None:
    for key in [
        "authenticated",
        "extraction_results",
        "extraction_failures",
        "categorization_result",
        "sheets-endpoint",
        "sheets-spreadsheet",
        "sheets-tab",
        "sheets-secret",
    ]:
        st.session_state.pop(key, None)
    st.query_params["view"] = "home"


st.sidebar.markdown(
    '<div class="rfs-brand"><span class="rfs-logo">R</span>'
    "<span>Raman Financial Services</span></div>",
    unsafe_allow_html=True,
)
st.sidebar.markdown('<div class="rfs-nav-label">Accounting tools</div>', unsafe_allow_html=True)
selected_page = st.sidebar.radio(
    "Accounting tools",
    NAV_ITEMS,
    key="main_nav",
    label_visibility="collapsed",
)
st.sidebar.markdown(
    '<div class="rfs-help"><strong>Secure workspace</strong><br>'
    "Review generated records before filing, remitting, lending, or investment decisions.</div>",
    unsafe_allow_html=True,
)
st.sidebar.markdown("[Public home](?view=home)")
st.sidebar.button("Sign out", on_click=sign_out, use_container_width=True)

st.markdown(
    '<div class="rfs-topbar"><div class="rfs-topbrand">'
    '<span class="rfs-logo">R</span><span>Raman Financial Services</span></div>'
    '<span class="rfs-status">Secure session</span></div>',
    unsafe_allow_html=True,
)
st.markdown(
    f'<div class="rfs-pagehead"><h1>{selected_page}</h1>'
    f"<p>{PAGE_COPY[selected_page]}</p></div>",
    unsafe_allow_html=True,
)

if selected_page == "Dashboard":
    st.markdown('<span class="rfs-kicker">Workspace</span>', unsafe_allow_html=True)
    overview_cols = st.columns(3)
    overview_cols[0].metric("Accounting workspaces", "7")
    overview_cols[1].metric("Supported statement sources", "8")
    overview_cols[2].metric("Output formats", "Excel, PDF, Word")

    st.subheader("Tools")
    tool_cards = [
        (
            "Extract statements",
            "Bank transactions",
            "Convert monthly bank and credit-card statements into reviewable transactions.",
        ),
        (
            "Auto categorize",
            "Bookkeeping categories",
            "Categorize existing transaction files with reusable merchant and direction rules.",
        ),
        (
            "Build annual file",
            "Annual bookkeeping",
            "Combine verified monthly workbooks without changing the statement order.",
        ),
        (
            "Payroll template",
            "Payroll records",
            "Calculate a pay period and add it to the employee's annual register.",
        ),
        (
            "Financial statements",
            "Compilation",
            "Create draft statements from searchable Schedule 100 and Schedule 125 PDFs.",
        ),
        (
            "Maximum mortgage",
            "Mortgage planning",
            "Estimate maximum mortgage capacity using the GDS calculation.",
        ),
        (
            "Real estate agent",
            "Investment analysis",
            "Rank property listings using financing, rent, comparable, and location inputs.",
        ),
    ]
    for row_start in range(0, len(tool_cards), 3):
        card_columns = st.columns(3)
        for column, (page, category, copy) in zip(card_columns, tool_cards[row_start : row_start + 3]):
            with column:
                with st.container(border=True):
                    st.markdown(f'<span class="rfs-kicker">{category}</span>', unsafe_allow_html=True)
                    st.markdown(f'<div class="rfs-card-title">{page}</div>', unsafe_allow_html=True)
                    st.markdown(f'<div class="rfs-card-copy">{copy}</div>', unsafe_allow_html=True)
                    st.button(
                        "Open tool",
                        key=f"dashboard_{page}",
                        on_click=go_to_page,
                        args=(page,),
                        use_container_width=True,
                    )

    with st.container(border=True):
        st.subheader("Review workflow")
        review_cols = st.columns(3)
        review_cols[0].markdown("**1. Process**\n\nUpload or enter the source records.")
        review_cols[1].markdown("**2. Reconcile**\n\nCompare results with the original documents.")
        review_cols[2].markdown("**3. Export**\n\nDownload the completed working file.")

if selected_page == "Extract statements":
    left, right = st.columns([1, 2])
    with left:
        selected_bank = st.selectbox("Bank", SUPPORTED_BANKS)
        uploaded_files = st.file_uploader(
            "Upload statement files",
            type=["pdf", "json"],
            accept_multiple_files=True,
            help="Upload one statement or several monthly statements.",
        )
        process_clicked = st.button("Extract transactions", type="primary", use_container_width=True)
        st.caption(
            "Tuned: BMO, CIBC, RBC bank accounts, RBC Visa Business, Tangerine, "
            "Vancity, and scanned debit/credit tables. Large PDFs are processed page by page."
        )

    with right:
        st.subheader("Processing status")
        status_placeholder = st.empty()
        if not uploaded_files:
            status_placeholder.info("Upload PDF or Docling JSON statements to begin.")

    if process_clicked:
        if not uploaded_files:
            st.error("Upload at least one statement.")
        else:
            successes = []
            failures = []
            progress = st.progress(0, text="Starting extraction...")
            for index, uploaded_file in enumerate(uploaded_files, start=1):
                try:
                    result = process_statement(uploaded_file, selected_bank)
                    successes.append(result)
                except Exception as exc:
                    failures.append({"file": uploaded_file.name, "error": str(exc)})
                progress.progress(index / len(uploaded_files), text=f"Processed {index} of {len(uploaded_files)}")
            progress.empty()
            st.session_state["extraction_results"] = successes
            st.session_state["extraction_failures"] = failures

    successes = st.session_state.get("extraction_results", [])
    failures = st.session_state.get("extraction_failures", [])
    if successes:
        st.success(f"Created {len(successes)} Excel file(s).")

        with st.expander("Send to Google Sheets", expanded=False):
            st.caption(
                "Send the extracted transaction rows directly to a private Google Sheet. "
                "Repeated uploads of the same statement are blocked by the connector."
            )
            connection_cols = st.columns(2)
            endpoint_url = connection_cols[0].text_input(
                "Apps Script Web App URL",
                value=os.getenv("GOOGLE_SHEETS_WEB_APP_URL", ""),
                placeholder="https://script.google.com/macros/s/.../exec",
                key="sheets-endpoint",
            )
            spreadsheet_target = connection_cols[1].text_input(
                "Google Sheet URL or spreadsheet ID",
                key="sheets-spreadsheet",
            )
            destination_cols = st.columns(2)
            destination_tab = destination_cols[0].text_input(
                "Destination tab",
                value="Bank Transactions",
                key="sheets-tab",
            )
            connector_secret = destination_cols[1].text_input(
                "Connector secret",
                value=os.getenv("GOOGLE_SHEETS_SHARED_SECRET", ""),
                type="password",
                key="sheets-secret",
            )
            setup_cols = st.columns([1, 1])
            setup_cols[0].download_button(
                "Download Apps Script connector",
                data=(APP_DIR / "google_sheets_connector.gs").read_bytes(),
                file_name="google_sheets_connector.gs",
                mime="text/plain",
                use_container_width=True,
            )
            send_clicked = setup_cols[1].button(
                "Send all to Google Sheets",
                type="primary",
                use_container_width=True,
            )
            with st.expander("Connector setup steps"):
                st.markdown(
                    "1. Create an Apps Script project at **script.google.com**.\n"
                    "2. Paste the downloaded connector code into `Code.gs`.\n"
                    "3. In **Project Settings → Script properties**, add `RFS_SHARED_SECRET` "
                    "and enter a long private value.\n"
                    "4. Select **Deploy → New deployment → Web app**. Execute it as yourself "
                    "and permit access required for the Streamlit server.\n"
                    "5. Paste the deployment URL ending in `/exec` and the same secret above."
                )
            st.caption(
                "The destination spreadsheet remains in its owner's Google Drive. "
                "The connector secret is used only for this browser session."
            )

            if send_clicked:
                upload_results = []
                with st.spinner("Sending transactions to Google Sheets..."):
                    for result in successes:
                        try:
                            response = sheets_sender.send_transactions(
                                endpoint_url=endpoint_url,
                                shared_secret=connector_secret,
                                spreadsheet=spreadsheet_target,
                                sheet_name=destination_tab,
                                frame=result["transactions"],
                                source_file=result["source"],
                                bank=result["bank"],
                            )
                            upload_results.append(
                                {
                                    "File": result["source"],
                                    "Status": (
                                        "Already uploaded"
                                        if response.get("duplicate")
                                        else "Sent"
                                    ),
                                    "Rows": int(response.get("rows_added", 0)),
                                }
                            )
                        except Exception as exc:
                            upload_results.append(
                                {
                                    "File": result["source"],
                                    "Status": "Error",
                                    "Rows": 0,
                                    "Message": str(exc),
                                }
                            )
                if all(item["Status"] != "Error" for item in upload_results):
                    st.success("Google Sheets upload completed.")
                else:
                    st.error("Some statements could not be sent.")
                st.dataframe(pd.DataFrame(upload_results), use_container_width=True, hide_index=True)

        if len(successes) > 1:
            annual_data, annual_transactions, annual_summary, annual_filename = merge_extracted_results(
                successes,
                "Annual_transactions.xlsx",
            )
            st.download_button(
                "Download one annual workbook",
                data=annual_data,
                file_name=annual_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )
            with st.expander("Preview annual summary"):
                st.dataframe(annual_summary, use_container_width=True, hide_index=True)
            st.download_button(
                "Download monthly Excel files as ZIP",
                data=make_zip(successes),
                file_name="bank_statement_excels.zip",
                mime="application/zip",
            )

        for result in successes:
            with st.expander(f"{result['source']} - {result['bank']}", expanded=len(successes) == 1):
                metrics = result["metrics"]
                cols = st.columns(4)
                cols[0].metric("Transactions", int(metrics.get("Number of transactions", 0)))
                cols[1].metric("Debits", f"${metrics.get('Total Debits', 0):,.2f}")
                cols[2].metric("Credits", f"${metrics.get('Total Credits', 0):,.2f}")
                cols[3].metric("Closing balance", f"${metrics.get('Closing Balance', 0):,.2f}")
                reconciliation = result["reconciliation"]
                if reconciliation["status"] == "reconciled":
                    st.success(reconciliation["message"])
                else:
                    st.warning(reconciliation["message"])
                st.download_button(
                    "Download Excel",
                    data=result["output_bytes"],
                    file_name=result["output_name"],
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"download-{result['output_name']}",
                )
                st.dataframe(result["transactions"], use_container_width=True, hide_index=True)

    if failures:
        st.error(f"{len(failures)} file(s) need review.")
        st.dataframe(pd.DataFrame(failures), use_container_width=True, hide_index=True)

if selected_page == "Build annual file":
    st.subheader("Combine monthly Excel files")
    monthly_files = st.file_uploader(
        "Upload monthly transaction workbooks",
        type=["xlsx"],
        accept_multiple_files=True,
        key="annual-files",
    )
    annual_name = st.text_input("Annual output name", value="Annual_transactions.xlsx")
    if st.button("Build annual workbook", type="primary"):
        if not monthly_files:
            st.error("Upload at least one monthly Excel file.")
        else:
            try:
                data, transactions, summary, filename = merge_uploaded_excels(monthly_files, annual_name)
                st.success(f"Merged {len(monthly_files)} workbook(s).")
                st.download_button(
                    "Download annual Excel",
                    data=data,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                )
                st.dataframe(summary, use_container_width=True, hide_index=True)
                with st.expander("Preview annual transactions"):
                    st.dataframe(transactions, use_container_width=True, hide_index=True)
            except Exception as exc:
                st.error(str(exc))

if selected_page == "Auto categorize":
    st.subheader("Automatic transaction categorization")
    st.caption(
        "Upload an Excel or CSV transaction file. Categories are assigned using local merchant "
        "and transaction-direction rules, then remain editable before download."
    )
    st.info(
        "Low-confidence and uncategorized transactions are marked for review. "
        "No transaction descriptions are sent to an external AI service."
    )

    upload_cols = st.columns([2, 1])
    transaction_file = upload_cols[0].file_uploader(
        "Transaction file",
        type=["xlsx", "csv"],
        key="categorizer-transactions",
        help="Supports Amount columns and separate Debit/Credit columns.",
    )
    custom_rule_file = upload_cols[1].file_uploader(
        "Custom rule CSV (optional)",
        type=["csv"],
        key="categorizer-rules",
        help="Columns: Keyword, Category, Applies To.",
    )
    st.download_button(
        "Download custom rule template",
        data=categorizer.custom_rule_template(),
        file_name="Category_Rules_Template.csv",
        mime="text/csv",
    )
    overwrite_categories = st.checkbox(
        "Replace existing nonblank categories",
        value=False,
        help="Opening Balance and Closing Totals are always protected.",
    )

    if st.button("Categorize transactions", type="primary"):
        if transaction_file is None:
            st.error("Upload an Excel or CSV transaction file.")
        else:
            try:
                custom_rules = []
                if custom_rule_file is not None:
                    custom_frame = pd.read_csv(io.BytesIO(custom_rule_file.getvalue()))
                    custom_rules = categorizer.parse_custom_rules(custom_frame)
                transactions = categorizer.load_transactions(
                    transaction_file.getvalue(),
                    transaction_file.name,
                )
                categorized = categorizer.categorize_transactions(
                    transactions,
                    custom_rules=custom_rules,
                    overwrite_existing=overwrite_categories,
                )
                st.session_state["categorization_result"] = {
                    "frame": categorized,
                    "source": transaction_file.name,
                    "custom_rules": len(custom_rules),
                }
            except Exception as exc:
                st.error(str(exc))

    category_result = st.session_state.get("categorization_result")
    if category_result:
        categorized = category_result["frame"]
        review_mask = categorized["Category Confidence"].isin(["Low", "Review"])
        metric_cols = st.columns(4)
        metric_cols[0].metric("Transactions", len(categorized))
        metric_cols[1].metric(
            "Categorized",
            int((categorized["Category"] != "Uncategorized").sum()),
        )
        metric_cols[2].metric("Needs review", int(review_mask.sum()))
        metric_cols[3].metric("Custom rules", category_result["custom_rules"])

        category_options = sorted(
            {
                rule.category
                for rule in categorizer.DEFAULT_RULES
            }
            | set(categorized["Category"].dropna().astype(str))
            | {"Uncategorized"}
        )
        edited = st.data_editor(
            categorized,
            use_container_width=True,
            hide_index=True,
            disabled=[
                column
                for column in categorized.columns
                if column != "Category"
            ],
            column_config={
                "Category": st.column_config.SelectboxColumn(
                    "Category",
                    options=category_options,
                    required=True,
                ),
                "Amount": st.column_config.NumberColumn(
                    "Amount",
                    format="$%.2f",
                ),
            },
            key="categorization-editor",
        )

        summary = categorizer.category_summary(edited)
        with st.expander("Category summary", expanded=True):
            st.dataframe(
                summary,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Net Amount": st.column_config.NumberColumn(
                        "Net Amount",
                        format="$%.2f",
                    )
                },
            )
        safe_stem = safe_name(Path(category_result["source"]).stem) or "Transactions"
        st.download_button(
            "Download categorized Excel",
            data=categorizer.export_categorized_workbook(edited),
            file_name=f"{safe_stem}_Categorized.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

if selected_page == "Payroll template":
    st.subheader("Payroll calculator")
    st.caption("Enter one pay period, then download a payslip PDF and updated payroll register.")
    payroll_pdoc_script = APP_DIR / "payroll_pdoc_google_sheets.gs"
    if payroll_pdoc_script.exists():
        st.download_button(
            "Download Google Sheets CRA PDOC helper",
            data=payroll_pdoc_script.read_text(encoding="utf-8"),
            file_name="payroll_pdoc_google_sheets.gs",
            mime="text/plain",
        )

    register_file = st.file_uploader(
        "Upload existing payroll register Excel",
        type=["xlsx"],
        key="payroll-register",
        help="Optional. If omitted, a new register will be created.",
    )
    try:
        payroll_register = load_payroll_register(register_file)
        if register_file is not None:
            st.success(f"Loaded {len(payroll_register)} existing payroll row(s).")
    except Exception as exc:
        st.error(str(exc))
        payroll_register = pd.DataFrame(columns=PAYROLL_COLUMNS)

    with st.form("payroll-form"):
        st.markdown("**Company**")
        company_cols = st.columns(2)
        company_name = company_cols[0].text_input("Company name", value="Raman Tax & Accounting Inc.")
        company_address = company_cols[1].text_input("Company address", value="Surrey, BC")

        st.markdown("**Employee**")
        emp_cols = st.columns(3)
        employee_name = emp_cols[0].text_input("Employee name")
        employee_id = emp_cols[1].text_input("Employee ID")
        position = emp_cols[2].text_input("Position")

        st.markdown("**Pay period**")
        period_cols = st.columns(4)
        frequency = period_cols[0].selectbox("Pay frequency", list(PAY_PERIODS.keys()), index=1)
        pay_start = period_cols[1].date_input("Pay start", value=date.today())
        pay_end = period_cols[2].date_input("Pay end", value=date.today())
        pay_date = period_cols[3].date_input("Pay date", value=date.today())

        st.markdown("**Earnings**")
        earn_cols = st.columns(4)
        hours = earn_cols[0].number_input("Regular hours", min_value=0.0, value=0.0, step=0.5)
        rate = earn_cols[1].number_input("Hourly rate", min_value=0.0, value=0.0, step=0.5)
        overtime_hours = earn_cols[2].number_input("Overtime hours", min_value=0.0, value=0.0, step=0.5)
        overtime_rate = earn_cols[3].number_input("Overtime rate", min_value=0.0, value=0.0, step=0.5)
        earn_cols2 = st.columns(4)
        stat_pay = earn_cols2[0].number_input("Stat pay", min_value=0.0, value=0.0, step=10.0)
        vacation_pay = earn_cols2[1].number_input("Vacation pay paid", min_value=0.0, value=0.0, step=10.0)
        bonus = earn_cols2[2].number_input("Bonus/other taxable pay", min_value=0.0, value=0.0, step=10.0)
        reimbursements = earn_cols2[3].number_input("Reimbursements", min_value=0.0, value=0.0, step=10.0)
        earn_cols3 = st.columns(3)
        salary_amount = earn_cols3[0].number_input("Salary amount", min_value=0.0, value=0.0, step=10.0)
        sick_pay = earn_cols3[1].number_input("Sick pay", min_value=0.0, value=0.0, step=10.0)
        vac_accrual = earn_cols3[2].number_input("Vacation accrual", min_value=0.0, value=0.0, step=10.0)

        st.markdown("**Year-to-date caps**")
        ytd_cols = st.columns(3)
        ytd_cpp = ytd_cols[0].number_input("YTD CPP already deducted", min_value=0.0, value=0.0, step=10.0)
        ytd_cpp2 = ytd_cols[1].number_input("YTD CPP2 already deducted", min_value=0.0, value=0.0, step=10.0)
        ytd_ei = ytd_cols[2].number_input("YTD EI already deducted", min_value=0.0, value=0.0, step=10.0)

        st.markdown("**Manual adjustment**")
        other_deductions = st.number_input("Other deductions", min_value=0.0, value=0.0, step=10.0)

        st.markdown("**CRA PDOC result**")
        pdoc_file = st.file_uploader(
            "Upload CRA PDOC PDF to use official deductions",
            type=["pdf"],
            key="payroll-pdoc-pdf",
            help="Optional. Download the result PDF from CRA PDOC and upload it here so the payslip uses CRA's CPP, EI, tax, and net pay.",
        )
        submitted = st.form_submit_button("Calculate payroll", type="primary")

    if submitted:
        payroll_input = {
            "frequency": frequency,
            "hours": hours,
            "rate": rate,
            "employee_id": employee_id,
            "pay_start": pay_start,
            "pay_end": pay_end,
            "pay_date": pay_date,
            "salary_amount": salary_amount,
            "overtime_hours": overtime_hours,
            "overtime_rate": overtime_rate,
            "stat_pay": stat_pay,
            "sick_pay": sick_pay,
            "vacation_pay": vacation_pay,
            "vac_accrual": vac_accrual,
            "bonus": bonus,
            "reimbursements": reimbursements,
            "other_deductions": other_deductions,
        }
        ytd_before = payroll_ytd_before(payroll_register, employee_id)
        payroll_input["ytd_cpp"] = ytd_before["cpp"] + ytd_cpp
        payroll_input["ytd_cpp2"] = ytd_cpp2
        payroll_input["ytd_ei"] = ytd_before["ei"] + ytd_ei
        calc = calculate_payroll(payroll_input)
        pdoc_result = {}
        try:
            if pdoc_file is not None:
                pdoc_result = extract_pdoc_result(pdoc_file)
                calc = apply_pdoc_to_calc(calc, pdoc_result, payroll_input)
        except Exception as exc:
            st.error(str(exc))
            st.stop()
        register_row = make_payroll_register_row(payroll_input, calc, ytd_before)
        updated_register = pd.concat(
            [payroll_register, pd.DataFrame([register_row], columns=PAYROLL_COLUMNS)],
            ignore_index=True,
        )
        st.session_state["payroll_calc"] = {
            "company": {"name": company_name, "address": company_address},
            "employee": {"name": employee_name or "Employee", "id": employee_id, "position": position},
            "payroll": {
                **payroll_input,
                "pay_start": pay_start,
                "pay_end": pay_end,
                "pay_date": pay_date,
            },
            "calc": calc,
            "pdoc_result": pdoc_result,
            "updated_register": updated_register,
        }

    saved = st.session_state.get("payroll_calc")
    if saved:
        calc = saved["calc"]
        st.markdown("**Payroll result**")
        metric_cols = st.columns(4)
        metric_cols[0].metric("Gross pay", f"${calc['gross']:,.2f}")
        metric_cols[1].metric("Employee deductions", f"${calc['total_deductions']:,.2f}")
        metric_cols[2].metric("Net pay", f"${calc['net']:,.2f}")
        metric_cols[3].metric("Employer cost add-on", f"${calc['employer_cpp'] + calc['employer_ei']:,.2f}")

        result_df = pd.DataFrame(
            [
                {"Item": "Regular pay", "Amount": calc["regular_pay"]},
                {"Item": "Salary amount", "Amount": saved["payroll"]["salary_amount"]},
                {"Item": "Overtime pay", "Amount": calc["overtime_pay"]},
                {"Item": "Sick pay", "Amount": saved["payroll"]["sick_pay"]},
                {"Item": "Gross pay", "Amount": calc["gross"]},
                {"Item": "CPP", "Amount": -calc["cpp"]},
                {"Item": "EI", "Amount": -calc["ei"]},
                {"Item": "Federal tax", "Amount": -calc["tax_fed"]},
                {"Item": "BC tax", "Amount": -calc["tax_prov"]},
                {"Item": "Other deductions", "Amount": -calc["other_deductions"]},
                {"Item": "Reimbursements", "Amount": calc["reimbursements"]},
                {"Item": "Net pay", "Amount": calc["net"]},
                {"Item": "Employer CPP", "Amount": calc["employer_cpp"]},
                {"Item": "Employer EI", "Amount": calc["employer_ei"]},
            ]
        )
        st.dataframe(result_df, use_container_width=True, hide_index=True)
        if calc.get("pdoc_source"):
            st.success("CRA PDOC PDF applied. Payslip and payroll register are using the official PDOC deduction amounts.")
            if calc.get("pdoc_warning"):
                st.warning(calc["pdoc_warning"])

        pdf_payroll = {
            **saved["payroll"],
            "cpp": calc["cpp"],
            "ei": calc["ei"],
            "tax_fed": calc["tax_fed"],
            "tax_prov": calc["tax_prov"],
            "total_deductions": calc["total_deductions"],
            "net": calc["net"],
        }
        pdf_bytes = build_payslip_pdf(saved["company"], saved["employee"], pdf_payroll, calc)
        st.download_button(
            "Download payslip PDF",
            data=pdf_bytes,
            file_name=f"{safe_name(saved['employee']['name']) or 'Employee'}_{saved['payroll']['pay_date']}_payslip.pdf",
            mime="application/pdf",
            type="primary",
        )
        st.download_button(
            "Download payslip Excel",
            data=build_payslip_excel(saved["company"], saved["employee"], pdf_payroll, calc),
            file_name=f"{safe_name(saved['employee']['name']) or 'Employee'}_{saved['payroll']['pay_date']}_payslip.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.download_button(
            "Download PD7A summary Excel",
            data=build_pd7a_excel(saved["company"], saved["payroll"], calc),
            file_name=f"PD7A_Remittance_{saved['payroll']['pay_date']}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.download_button(
            "Download PD7A summary PDF",
            data=build_pd7a_pdf(saved["company"], saved["payroll"], calc),
            file_name=f"PD7A_Remittance_{saved['payroll']['pay_date']}.pdf",
            mime="application/pdf",
        )
        register_bytes = export_payroll_register(saved["updated_register"])
        st.download_button(
            "Download updated payroll register",
            data=register_bytes,
            file_name="Payroll_Register_Updated.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        with st.expander("Preview payroll register row"):
            st.dataframe(saved["updated_register"].tail(1), use_container_width=True, hide_index=True)
        if calc.get("pdoc_source"):
            st.info("Keep the CRA PDOC PDF with the payroll file as calculation support.")
        else:
            st.warning("Payroll calculations are estimates. Upload a CRA PDOC result PDF to make the payslip match PDOC.")

if selected_page == "Financial statements":
    st.subheader("Compiled financial statements")
    st.caption(
        "Upload one combined S100/S125 PDF or two separate schedule PDFs. The app extracts current-year GIFI amounts and creates "
        "a draft Compilation Engagement Report package (formerly called Notice to Reader)."
    )
    st.warning(
        "The generated package is a draft. A qualified practitioner must review the classifications, "
        "basis of accounting, CSRS 4200 report wording, report date, and signature before issuance."
    )

    with st.form("financial-statements-form"):
        combined_schedule_file = st.file_uploader(
            "Combined Schedule 100 and Schedule 125 PDF",
            type=["pdf"],
            key="combined-schedule-100-125",
            help="Upload the PDF package containing both the balance sheet and income statement GIFI schedules.",
        )
        with st.expander("Or upload Schedule 100 and Schedule 125 separately"):
            file_cols = st.columns(2)
            schedule_100_file = file_cols[0].file_uploader(
                "Schedule 100 - Balance Sheet",
                type=["pdf"],
                key="schedule-100",
            )
            schedule_125_file = file_cols[1].file_uploader(
                "Schedule 125 - Income Statement",
                type=["pdf"],
                key="schedule-125",
            )

        st.markdown("**Corporation and reporting period**")
        fs_auto_details = st.checkbox(
            "Read corporation name, business number, and fiscal year end from the PDF",
            value=True,
        )
        company_cols = st.columns(3)
        fs_company_name = company_cols[0].text_input(
            "Corporation legal name",
            disabled=fs_auto_details,
        )
        fs_business_number = company_cols[1].text_input(
            "Business number (optional)",
            disabled=fs_auto_details,
        )
        fs_year_end = company_cols[2].date_input(
            "Fiscal year end",
            value=date.today(),
            disabled=fs_auto_details,
        )

        st.markdown("**Practitioner information**")
        firm_cols = st.columns(3)
        fs_firm_name = firm_cols[0].text_input("Accounting firm", value="Raman Financial Services")
        fs_firm_address = firm_cols[1].text_input("Practitioner address", value="Surrey, BC")
        fs_report_date = firm_cols[2].date_input("Report date", value=date.today())

        fs_basis = st.text_area(
            "Basis of accounting note",
            value=(
                "The financial information has been prepared using a basis of accounting selected by "
                "management. Revenue and expenses are recorded using the accrual method, capital assets "
                "are recorded at cost less accumulated amortization, and income taxes are recorded on "
                "the basis used in the corporation's income tax return."
            ),
            height=120,
        )
        with st.expander("Compilation Engagement Report wording"):
            fs_report_text = st.text_area(
                "Approved report wording",
                value=fs_generator.DEFAULT_COMPILATION_REPORT,
                height=300,
                help=(
                    "Placeholders {company_name} and {year_end} are filled automatically. "
                    "Replace this draft with your firm's approved CSRS 4200 wording."
                ),
            )

        fs_submitted = st.form_submit_button("Create draft financial statements", type="primary")

    if fs_submitted:
        if combined_schedule_file is None and (
            schedule_100_file is None or schedule_125_file is None
        ):
            st.error("Upload one combined S100/S125 PDF, or upload both schedules separately.")
        else:
            try:
                with st.spinner("Extracting GIFI data and preparing the report package..."):
                    if combined_schedule_file is not None:
                        schedule_100, schedule_125, detected = (
                            fs_generator.extract_combined_schedule_pdf(
                                combined_schedule_file.getvalue()
                            )
                        )
                    else:
                        schedule_100 = fs_generator.extract_schedule_pdf(
                            schedule_100_file.getvalue(), "100"
                        )
                        schedule_125 = fs_generator.extract_schedule_pdf(
                            schedule_125_file.getvalue(), "125"
                        )
                        detected_100 = fs_generator.extract_statement_metadata(
                            schedule_100_file.getvalue()
                        )
                        detected_125 = fs_generator.extract_statement_metadata(
                            schedule_125_file.getvalue()
                        )
                        detected = {
                            key: detected_100.get(key) or detected_125.get(key)
                            for key in ("company_name", "business_number", "year_end")
                        }

                    company_name = (
                        detected.get("company_name", "") if fs_auto_details else fs_company_name.strip()
                    )
                    business_number = (
                        detected.get("business_number", "") if fs_auto_details else fs_business_number.strip()
                    )
                    year_end = (
                        detected.get("year_end") if fs_auto_details else fs_year_end
                    )
                    if not company_name:
                        raise RuntimeError(
                            "The corporation name was not found. Turn off automatic corporation details and enter it manually."
                        )
                    if year_end is None:
                        raise RuntimeError(
                            "The fiscal year end was not found. Turn off automatic corporation details and enter it manually."
                        )

                    metadata = {
                        "company_name": company_name,
                        "business_number": business_number,
                        "year_end": year_end,
                        "firm_name": fs_firm_name.strip() or "Accounting practitioner",
                        "firm_address": fs_firm_address.strip(),
                        "report_date": fs_report_date,
                        "basis_of_accounting": fs_basis.strip(),
                        "report_text": fs_report_text.strip(),
                    }
                    pdf_bytes = fs_generator.build_financial_statement_pdf(
                        schedule_100.entries, schedule_125.entries, metadata
                    )
                    docx_bytes = fs_generator.build_financial_statement_docx(
                        schedule_100.entries, schedule_125.entries, metadata
                    )
                    st.session_state["financial_statement_result"] = {
                        "balance": schedule_100.entries,
                        "income": schedule_125.entries,
                        "warnings": schedule_100.warnings + schedule_125.warnings,
                        "metadata": metadata,
                        "pdf": pdf_bytes,
                        "docx": docx_bytes,
                        "pdf_name": fs_generator.suggested_filename(company_name, year_end, "pdf"),
                        "docx_name": fs_generator.suggested_filename(company_name, year_end, "docx"),
                    }
            except Exception as exc:
                st.error(str(exc))

    fs_result = st.session_state.get("financial_statement_result")
    if fs_result:
        metadata = fs_result.get("metadata", {})
        if metadata:
            st.info(
                f"Prepared for {metadata.get('company_name', '')} | "
                f"Business number: {metadata.get('business_number') or 'Not found'} | "
                f"Year end: {metadata.get('year_end')}"
            )
        if fs_result["warnings"]:
            st.error("Review required before issuance:")
            for warning in fs_result["warnings"]:
                st.write(f"- {warning}")
        else:
            st.success("Schedule 100 balances and Schedule 125 reconciles.")

        download_cols = st.columns(2)
        download_cols[0].download_button(
            "Download draft PDF",
            data=fs_result["pdf"],
            file_name=fs_result["pdf_name"],
            mime="application/pdf",
            type="primary",
            use_container_width=True,
        )
        download_cols[1].download_button(
            "Download editable Word file",
            data=fs_result["docx"],
            file_name=fs_result["docx_name"],
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )

        preview_cols = st.columns(2)
        with preview_cols[0]:
            st.markdown("**Schedule 100 extracted data**")
            st.dataframe(fs_result["balance"], use_container_width=True, hide_index=True)
        with preview_cols[1]:
            st.markdown("**Schedule 125 extracted data**")
            st.dataframe(fs_result["income"], use_container_width=True, hide_index=True)

if selected_page == "Maximum mortgage":
    st.subheader("Maximum Mortgage Under GDSR")
    st.caption("Estimate mortgage capacity using gross income and housing costs.")

    with st.form("gds_mortgage_form"):
        income_col, rate_col, amortization_col = st.columns(3)
        annual_income = income_col.number_input(
            "Annual household income",
            min_value=0.0,
            value=120000.0,
            step=5000.0,
            format="%.2f",
        )
        contract_rate = rate_col.number_input(
            "Mortgage contract rate (%)",
            min_value=0.0,
            value=4.50,
            step=0.05,
            format="%.2f",
        )
        amortization = amortization_col.selectbox("Amortization", [25, 30], format_func=lambda value: f"{value} years")

        tax_col, heat_col, condo_col = st.columns(3)
        annual_taxes = tax_col.number_input(
            "Annual property taxes",
            min_value=0.0,
            value=4800.0,
            step=100.0,
            format="%.2f",
        )
        monthly_heat = heat_col.number_input(
            "Monthly heating cost",
            min_value=0.0,
            value=150.0,
            step=10.0,
            format="%.2f",
        )
        monthly_condo = condo_col.number_input(
            "Monthly condo fees",
            min_value=0.0,
            value=0.0,
            step=25.0,
            format="%.2f",
        )

        gds_col, down_col = st.columns(2)
        gds_limit = gds_col.number_input(
            "Maximum GDS ratio (%)",
            min_value=1.0,
            max_value=100.0,
            value=39.0,
            step=0.5,
        )
        down_payment = down_col.number_input(
            "Down payment for purchase-price estimate (%)",
            min_value=0.0,
            max_value=99.0,
            value=20.0,
            step=1.0,
        )
        mortgage_clicked = st.form_submit_button(
            "Calculate maximum mortgage",
            type="primary",
            use_container_width=True,
        )

    if mortgage_clicked:
        try:
            st.session_state["gds_result"] = mortgage.calculate_maximum_mortgage(
                annual_household_income=annual_income,
                contract_rate_pct=contract_rate,
                amortization_years=amortization,
                annual_property_taxes=annual_taxes,
                monthly_heating=monthly_heat,
                monthly_condo_fees=monthly_condo,
                gds_limit_pct=gds_limit,
                down_payment_pct=down_payment,
            )
        except ValueError as exc:
            st.error(str(exc))

    gds_result = st.session_state.get("gds_result")
    if gds_result:
        result_cols = st.columns(4)
        result_cols[0].metric("Maximum mortgage", f"${gds_result.maximum_mortgage:,.0f}")
        result_cols[1].metric("Maximum mortgage payment", f"${gds_result.maximum_mortgage_payment:,.2f}/month")
        result_cols[2].metric("Qualifying rate", f"{gds_result.qualifying_rate:.2%}")
        result_cols[3].metric("Estimated purchase price", f"${gds_result.estimated_purchase_price:,.0f}")

        st.markdown("**GDS calculation**")
        st.dataframe(
            pd.DataFrame(
                [
                    ["Gross monthly income", gds_result.gross_monthly_income],
                    ["Maximum housing cost under GDS", gds_result.maximum_housing_cost],
                    ["Less: monthly property taxes", -gds_result.property_tax_monthly],
                    ["Less: monthly heating", -gds_result.heating_monthly],
                    ["Less: 50% of condo fees", -gds_result.condo_fee_portion],
                    ["Available mortgage payment", gds_result.maximum_mortgage_payment],
                ],
                columns=["Calculation", "Monthly amount"],
            ).style.format({"Monthly amount": "${:,.2f}"}),
            use_container_width=True,
            hide_index=True,
        )
        st.code(
            "Maximum mortgage payment = (gross monthly income x GDS limit) "
            "- property taxes - heating - 50% of condo fees",
            language="text",
        )
        if gds_result.maximum_mortgage <= 0:
            st.warning("The entered housing costs use all available GDS capacity.")

    st.info(
        "This is a planning estimate. It uses the greater of the contract rate plus 2% or 5.25%. "
        "A lender will also review TDS, credit, income, down payment, property type, and its own policies."
    )

if selected_page == "Real estate agent":
    st.subheader("Real Estate Investment Agent")
    st.caption("Rank listings using financing costs, rents, comparable sales, and location signals.")

    with st.form("real_estate_search_form"):
        source_mode = st.radio(
            "Listing source",
            ["Realtor.ca map search", "Saved listing CSV"],
            horizontal=True,
        )
        realtor_url = st.text_area(
            "Realtor.ca map search URL",
            value=DEFAULT_REALTOR_URL,
            height=90,
            disabled=source_mode != "Realtor.ca map search",
        )
        listing_csv = st.file_uploader(
            "Realtor listing CSV",
            type=["csv"],
            disabled=source_mode != "Saved listing CSV",
        )

        st.markdown("**Financing assumptions**")
        rate_input, amortization_input, down_input = st.columns(3)
        investment_rate = rate_input.number_input(
            "Mortgage rate (%)",
            min_value=0.0,
            value=5.25,
            step=0.05,
        )
        investment_amortization = amortization_input.selectbox(
            "Mortgage amortization",
            [25, 30],
            format_func=lambda value: f"{value} years",
        )
        investment_down = down_input.number_input(
            "Down payment (%)",
            min_value=0.0,
            max_value=99.0,
            value=20.0,
            step=1.0,
        )
        buffer_input, vacancy_input = st.columns(2)
        monthly_buffer = buffer_input.number_input(
            "Monthly repair/insurance buffer",
            min_value=0.0,
            value=250.0,
            step=25.0,
        )
        vacancy_rate = vacancy_input.number_input(
            "Vacancy allowance (%)",
            min_value=0.0,
            max_value=100.0,
            value=3.0,
            step=0.5,
        )

        with st.expander("Optional enrichment files"):
            enrich_cols = st.columns(3)
            zealty_upload = enrich_cols[0].file_uploader("Comparable sales JSON", type=["json"])
            rental_upload = enrich_cols[1].file_uploader("Rental estimates JSON", type=["json"])
            signal_upload = enrich_cols[2].file_uploader("Transit, school and development JSON", type=["json"])

        investment_clicked = st.form_submit_button(
            "Run investment search",
            type="primary",
            use_container_width=True,
        )

    if investment_clicked:
        assumptions = InvestmentAssumptions(
            mortgage_interest_rate=investment_rate / 100,
            amortization_years=investment_amortization,
            down_payment_pct=investment_down / 100,
            monthly_buffer=monthly_buffer,
            vacancy_pct=vacancy_rate / 100,
        )
        try:
            with st.spinner("Collecting and ranking properties..."):
                investment_rows, investment_counts = run_real_estate_search(
                    source_mode,
                    realtor_url,
                    listing_csv,
                    zealty_upload,
                    rental_upload,
                    signal_upload,
                    assumptions,
                )
            if not investment_rows:
                st.warning("No supported listings were found in the selected search.")
            else:
                st.session_state["investment_rows"] = investment_rows
                st.session_state["investment_counts"] = investment_counts
        except Exception as exc:
            st.error(str(exc))

    investment_rows = st.session_state.get("investment_rows", [])
    if investment_rows:
        investment_df = pd.DataFrame(
            [{key: value for key, value in row.items() if not key.startswith("_")} for row in investment_rows]
        )
        investment_top = investment_df.sort_values("Investment Score", ascending=False).head(10)
        price_per_sqft = investment_df["Price / Sq Ft"].dropna()
        investment_metrics = st.columns(3)
        investment_metrics[0].metric("Listings found", len(investment_df))
        investment_metrics[1].metric("Average list price", f"${investment_df['List Price'].mean():,.0f}")
        investment_metrics[2].metric(
            "Median price per sq. ft.",
            f"${price_per_sqft.median():,.0f}" if not price_per_sqft.empty else "Not available",
        )

        st.markdown("**Top investment opportunities**")
        st.dataframe(
            investment_top[
                [
                    "Address",
                    "City",
                    "List Price",
                    "Estimated Rent",
                    "Estimated Cash Flow",
                    "Investment Score",
                    "Flags",
                    "Notes",
                ]
            ],
            use_container_width=True,
            hide_index=True,
        )
        with st.expander("Full property database"):
            st.dataframe(investment_df, use_container_width=True, hide_index=True)
        st.download_button(
            "Download investment workbook",
            data=real_estate_excel_bytes(investment_df, investment_top),
            file_name="Real_Estate_Investment_Analysis.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
    else:
        st.info("Use a Realtor.ca map URL or upload a saved listing CSV to begin.")

if selected_page == "Guide":
    st.subheader("Recommended file names")
    st.caption("You may use any filename. This format makes monthly and annual files easier to sort.")
    st.code(
        "2025-02_BMO_Chequing.pdf\n"
        "2025-02_CIBC_Chequing.pdf\n"
        "2025-02_RBC_Chequing.pdf\n"
        "2025-02_RBC_Visa.pdf\n"
        "2025-02_TD_Chequing.pdf\n"
        "2025-02_Tangerine_Chequing.pdf\n"
        "2025-02_Vancity_Chequing.pdf",
        language="text",
    )
    st.subheader("Review checklist")
    st.markdown(
        "1. Upload statement files and extract Excel.\n"
        "2. Compare total debits, credits, opening balance and closing balance with the statement.\n"
        "3. Keep original PDFs until the annual accounting file is complete.\n"
        "4. Merge the verified monthly workbooks into an annual workbook."
    )
    st.subheader("Automatic categorization")
    st.markdown(
        "1. Open **Auto categorize** and upload an extracted Excel workbook or transaction CSV.\n"
        "2. Keep **Replace existing nonblank categories** off when you want to preserve reviewed work.\n"
        "3. Optionally upload a custom rule CSV for recurring merchants or customers.\n"
        "4. Review all rows marked **Low** or **Review** and change the Category directly in the table.\n"
        "5. Download the categorized workbook with its Category Summary and Built-in Rules sheets."
    )
    st.subheader("Send to Google Sheets")
    st.markdown(
        "1. Extract and review the statement transactions.\n"
        "2. Open **Send to Google Sheets** and download the Apps Script connector.\n"
        "3. Deploy the connector from your own Google account and create its private secret.\n"
        "4. Enter the connector URL, destination spreadsheet, tab name and secret.\n"
        "5. Select **Send all to Google Sheets**. Repeating the same upload will not duplicate it."
    )
    st.subheader("Financial statement files")
    st.markdown(
        "1. Export **Schedule 100** and **Schedule 125** from the T2 software as one combined PDF or two separate PDFs.\n"
        "2. Open **Financial statements** and upload the combined file. Use the separate boxes only when the schedules are separate.\n"
        "3. The corporation name, business number, and fiscal year end are read automatically. Enter the practitioner, basis of accounting, and report date.\n"
        "4. Confirm the balance and income reconciliation messages.\n"
        "5. Review and sign the editable Word draft before providing it to a client or third party."
    )
    st.code(
        "2025-12-31_CompanyName_S100_S125.pdf\n"
        "or\n"
        "2025-12-31_CompanyName_S100.pdf\n"
        "2025-12-31_CompanyName_S125.pdf",
        language="text",
    )
    st.subheader("Mortgage and real estate tools")
    st.markdown(
        "1. Use **Maximum mortgage** for a GDS-based planning estimate.\n"
        "2. Enter heating and condo fees separately; the calculator includes 50% of condo fees.\n"
        "3. Use **Real estate agent** with a Realtor.ca map search URL or a saved listing CSV.\n"
        "4. Add optional comparable-sale, rental, and location JSON files for a stronger score.\n"
        "5. Review financing, rent, expenses, title, zoning, condition, and lender approval before acting."
    )
    st.info("Files uploaded to this app are processed for the current session. Configure your hosting provider's privacy and retention settings before using real client statements.")
