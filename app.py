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
from pathlib import Path

import pandas as pd
import streamlit as st


APP_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(APP_DIR))

import bmo_docling_to_excel as extractor

importlib.reload(extractor)


st.set_page_config(
    page_title="Bank Statement Extractor",
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

    st.title("Bank Statement Extractor")
    password = st.text_input("Password", type="password")
    if st.button("Sign in", type="primary"):
        if hmac.compare_digest(password, expected):
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("Incorrect password.")
    st.stop()


require_password()


SUPPORTED_BANKS = ["Auto-detect", "BMO", "CIBC", "RBC", "Tangerine", "TD", "Other bank"]


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

    calculated = round(
        float(opening.iloc[0])
        - float(pd.to_numeric(normal["Debit"], errors="coerce").fillna(0).sum())
        + float(pd.to_numeric(normal["Credit"], errors="coerce").fillna(0).sum()),
        2,
    )
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


st.title("Bank Statement Extractor")
st.caption("Convert bank statements into reviewable Excel transactions, then combine monthly files into an annual workbook.")

extract_tab, annual_tab, guide_tab = st.tabs(["Extract statements", "Build annual file", "Guide"])

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
        st.caption("Tuned: BMO, CIBC, RBC and Tangerine. TD and other banks use the Docling fallback.")

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
                    st.download_button(
                        "Download all Excel files",
                        data=make_zip(successes),
                        file_name="bank_statement_excels.zip",
                        mime="application/zip",
                        type="primary",
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

with guide_tab:
    st.subheader("Recommended file names")
    st.code(
        "2026-01_BMO_Chequing.pdf\n"
        "2026-01_CIBC_Chequing.pdf\n"
        "2026-01_RBC_Chequing.pdf\n"
        "2026-01_TD_Chequing.pdf\n"
        "2026-01_Tangerine_Chequing.pdf",
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
