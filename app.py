from __future__ import annotations

import importlib
import hmac
import io
import json
import os
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
import real_estate_research_agent as real_estate

importlib.reload(extractor)
importlib.reload(real_estate)


st.set_page_config(
    page_title="Raman Financial Services - Bank Statement Extractor",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
    .block-container {max-width: 1120px; padding-top: 2rem; padding-bottom: 4rem;}
    h1, h2, h3 {letter-spacing: 0;}
    [data-testid="stMetric"] {border: 1px solid #d9dee7; padding: 14px; border-radius: 6px;}
    .status-ok {padding: 12px 14px; border-left: 4px solid #17864b; background: #eff8f3;}
    .status-error {padding: 12px 14px; border-left: 4px solid #c73838; background: #fff2f2;}
    </style>
    """,
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

    st.title("Raman Financial Services")
    st.subheader("Bank Statement Extractor")
    st.markdown("[ramanfinancialservices.ca](https://ramanfinancialservices.ca/)")
    password = st.text_input("Password", type="password")
    if st.button("Sign in", type="primary"):
        if hmac.compare_digest(password, expected):
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("Incorrect password.")
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


def safe_name(value: str) -> str:
    return "".join(char if char.isalnum() or char in "-_. " else "_" for char in value).strip()


def docling_json_for_pdf(pdf_path: Path, work_dir: Path) -> Path:
    try:
        from docling.document_converter import DocumentConverter
    except ImportError as exc:
        raise RuntimeError(
            "BMO PDF conversion requires Docling. Install the website requirements and restart the app."
        ) from exc

    json_path = work_dir / f"{pdf_path.stem}.docling.json"
    result = DocumentConverter().convert(str(pdf_path))
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
    if input_path.suffix.lower() == ".pdf" and bank not in {"BMO", "TD", "Other bank"}:
        try:
            rows = extractor.extract_from_statement_file(input_path, bank)
        except RuntimeError as exc:
            direct_error = exc

    if not rows:
        effective_path = input_path
        if input_path.suffix.lower() == ".pdf":
            effective_path = docling_json_for_pdf(input_path, work_dir)
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


DEFAULT_REALTOR_URL = (
    "https://www.realtor.ca/map#ZoomLevel=15&Center=49.158114%2C-122.846239"
    "&LatitudeMax=49.16707&LongitudeMax=-122.81311&LatitudeMin=49.14916"
    "&LongitudeMin=-122.87937&Sort=6-D&PropertyTypeGroupID=1"
    "&TransactionTypeId=2&PropertySearchTypeId=1&OwnershipTypeGroupId=1&Currency=CAD"
)


def uploaded_json_temp(uploaded_file) -> Path | None:
    if uploaded_file is None:
        return None
    payload = uploaded_file.getvalue()
    json.loads(payload.decode("utf-8"))
    temp = tempfile.NamedTemporaryFile(delete=False, suffix=".json")
    temp.write(payload)
    temp.close()
    return Path(temp.name)


def real_estate_public_rows(rows: list[dict]) -> pd.DataFrame:
    return pd.DataFrame([{key: value for key, value in row.items() if not key.startswith("_")} for row in rows])


def real_estate_excel_bytes(tables: dict[str, pd.DataFrame]) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for sheet_name, df in tables.items():
            df.to_excel(writer, index=False, sheet_name=sheet_name[:31])
            worksheet = writer.sheets[sheet_name[:31]]
            for column_cells in worksheet.columns:
                width = max(len(str(cell.value or "")) for cell in column_cells)
                worksheet.column_dimensions[column_cells[0].column_letter].width = min(max(width + 2, 10), 44)
    return buffer.getvalue()


def uploaded_csv_temp(uploaded_file) -> Path | None:
    if uploaded_file is None:
        return None
    temp = tempfile.NamedTemporaryFile(delete=False, suffix=".csv")
    temp.write(uploaded_file.getvalue())
    temp.close()
    return Path(temp.name)


def run_real_estate_search(
    realtor_url: str,
    realtor_csv_file,
    zealty_file,
    rental_file,
    signal_file,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    realtor_csv_path = uploaded_csv_temp(realtor_csv_file)
    zealty_path = uploaded_json_temp(zealty_file)
    rental_path = uploaded_json_temp(rental_file)
    signal_path = uploaded_json_temp(signal_file)

    assumptions = real_estate.Assumptions()
    realtor = real_estate.RealtorSavedSearchProvider(
        csv_path=realtor_csv_path,
        search_url=None if realtor_csv_path else realtor_url,
    )
    zealty = real_estate.ZealtyProvider(zealty_path)
    rentals = real_estate.RentalProvider(rental_path)
    signals = real_estate.SignalProvider(signal_path)

    rows = []
    for listing in realtor.fetch():
        rows.append(
            real_estate.score_property(
                listing,
                zealty.fetch(listing),
                rentals.fetch(listing),
                signals.fetch(listing),
                assumptions,
            )
        )

    df = real_estate_public_rows(rows)
    top = df.sort_values("Investment Score", ascending=False).head(10) if not df.empty else df
    return df, top


st.title("Raman Financial Services")
st.subheader("Bank Statement Extractor")
st.markdown("[ramanfinancialservices.ca](https://ramanfinancialservices.ca/)")
st.caption("Convert bank statements into reviewable Excel transactions, then combine monthly files into an annual workbook.")

extract_tab, annual_tab, payroll_tab, real_estate_tab, guide_tab = st.tabs(
    ["Extract statements", "Build annual file", "Payroll template", "Real Estate Agent", "Guide"]
)

with extract_tab:
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
        st.caption("Tuned: BMO, CIBC, RBC bank accounts, RBC Visa Business and Tangerine. TD and other banks use the Docling fallback.")

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

            if successes:
                st.success(f"Created {len(successes)} Excel file(s).")
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

with annual_tab:
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

with payroll_tab:
    st.subheader("Payroll calculator")
    st.caption("Enter one pay period, then download a payslip PDF and updated payroll register.")

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
        register_bytes = export_payroll_register(saved["updated_register"])
        st.download_button(
            "Download updated payroll register",
            data=register_bytes,
            file_name="Payroll_Register_Updated.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        with st.expander("Preview payroll register row"):
            st.dataframe(saved["updated_register"].tail(1), use_container_width=True, hide_index=True)
        st.warning("Payroll calculations should be reviewed against CRA PDOC before remitting or filing.")

with real_estate_tab:
    st.subheader("Real Estate Investment Agent")
    st.caption("Paste a Realtor.ca map search URL, run the search, then review ranked investment opportunities.")
    st.info(
        "Use the Realtor.ca map-search link, not a single property page. "
        "The URL should start with https://www.realtor.ca/map# and include LatitudeMax / LatitudeMin. "
        "If Realtor.ca blocks the live search, upload a Realtor.ca CSV export below."
    )

    realtor_url = st.text_area("Realtor.ca map search URL", value=DEFAULT_REALTOR_URL, height=110)
    realtor_csv_file = st.file_uploader(
        "Realtor.ca CSV fallback",
        type=["csv"],
        help="Use this when Realtor.ca blocks the live map request. Columns should include Address, City, MLS Number, List Price, bedrooms, bathrooms, square feet, etc.",
    )

    with st.expander("Optional enrichment files"):
        enrich_cols = st.columns(3)
        zealty_file = enrich_cols[0].file_uploader("Zealty JSON", type=["json"], key="zealty-json")
        rental_file = enrich_cols[1].file_uploader("Rental JSON", type=["json"], key="rental-json")
        signal_file = enrich_cols[2].file_uploader(
            "Transit / school / development JSON",
            type=["json"],
            key="signal-json",
        )

    if st.button("Run real estate search", type="primary", use_container_width=True):
        if not realtor_url.strip() and realtor_csv_file is None:
            st.error("Paste a Realtor.ca map URL or upload a Realtor.ca CSV first.")
        elif realtor_csv_file is None and "/real-estate/" in realtor_url.lower():
            st.error(
                "That is a single property link. Open Realtor.ca map view, apply your filters, "
                "then copy the map URL from the address bar."
            )
        else:
            try:
                with st.spinner("Searching Realtor.ca and ranking properties..."):
                    property_df, top_df = run_real_estate_search(
                        realtor_url.strip(),
                        realtor_csv_file,
                        zealty_file,
                        rental_file,
                        signal_file,
                    )
                st.session_state["real_estate_property_df"] = property_df
                st.session_state["real_estate_top_df"] = top_df
            except Exception as exc:
                st.error(str(exc))

    property_df = st.session_state.get("real_estate_property_df")
    top_df = st.session_state.get("real_estate_top_df")

    if isinstance(property_df, pd.DataFrame) and not property_df.empty:
        metric_cols = st.columns(4)
        metric_cols[0].metric("Listings", f"{len(property_df):,}")
        metric_cols[1].metric("Average list price", f"${property_df['List Price'].mean():,.0f}")
        price_sqft = pd.to_numeric(property_df["Price / Sq Ft"], errors="coerce").dropna()
        metric_cols[2].metric("Median price / sq ft", f"${price_sqft.median():,.0f}" if not price_sqft.empty else "Pending")
        metric_cols[3].metric("Top score", f"{top_df['Investment Score'].max():.0f}" if not top_df.empty else "Pending")

        st.markdown("**Top 10 Investment Opportunities**")
        top_columns = [
            "Address",
            "City",
            "List Price",
            "1-Year Change %",
            "5-Year Change %",
            "Estimated Rent",
            "Estimated Cash Flow",
            "Investment Score",
            "Flags",
            "Notes",
        ]
        visible_top_columns = [column for column in top_columns if column in top_df.columns]
        st.dataframe(top_df[visible_top_columns], use_container_width=True, hide_index=True)

        st.markdown("**Property Database**")
        st.dataframe(property_df, use_container_width=True, hide_index=True)

        download_cols = st.columns(2)
        download_cols[0].download_button(
            "Download property CSV",
            data=property_df.to_csv(index=False).encode("utf-8-sig"),
            file_name="property_database_rows.csv",
            mime="text/csv",
            use_container_width=True,
        )
        download_cols[1].download_button(
            "Download Excel workbook",
            data=real_estate_excel_bytes(
                {
                    "Property Database": property_df,
                    "Top 10 Opportunities": top_df,
                }
            ),
            file_name="property_database.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    else:
        st.info("Run a Realtor.ca search to populate the property database.")

with guide_tab:
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
    st.info("Files uploaded to this app are processed for the current session. Configure your hosting provider's privacy and retention settings before using real client statements.")
