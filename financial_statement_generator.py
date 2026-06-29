from __future__ import annotations

import io
import re
import tempfile
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Iterable
from xml.sax.saxutils import escape

import pandas as pd


# Common GIFI labels are fallbacks. Most tax-software PDFs include the full
# description beside the code, and that description is preferred.
GIFI_LABELS = {
    1001: "Cash",
    1060: "Accounts receivable",
    1120: "Inventories",
    1180: "Short-term investments",
    1300: "Due from shareholder(s)/director(s)",
    1400: "Due from related parties",
    1484: "Prepaid expenses",
    1599: "Total current assets",
    1600: "Land",
    1680: "Buildings",
    1740: "Machinery and equipment",
    1787: "Accumulated amortization of machinery and equipment",
    2008: "Total tangible capital assets",
    2009: "Total accumulated amortization of tangible capital assets",
    2010: "Intangible assets",
    2178: "Total intangible capital assets",
    2179: "Total accumulated amortization of intangible capital assets",
    2180: "Long-term investments",
    2240: "Investments in related corporations",
    2420: "Other assets",
    2599: "Total assets",
    2600: "Bank overdraft",
    2621: "Trade payables",
    2680: "Taxes payable",
    2700: "Short-term debt",
    2780: "Due to shareholder(s)/director(s)",
    2840: "Due to related parties",
    2920: "Current portion of long-term liability",
    3139: "Total current liabilities",
    3140: "Long-term debt",
    3143: "Chartered bank loan",
    3260: "Due to shareholder(s)/director(s)",
    3450: "Total long-term liabilities",
    3499: "Total liabilities",
    3500: "Common shares",
    3520: "Preferred shares",
    3540: "Contributed and other surplus",
    3580: "Accumulated other comprehensive income",
    3600: "Retained earnings (deficit)",
    3620: "Total shareholder equity",
    3640: "Total liabilities and shareholder equity",
    8000: "Trade sales of goods and services",
    8090: "Investment revenue",
    8140: "Rental revenue",
    8230: "Other revenue",
    8299: "Total revenue",
    8300: "Opening inventory",
    8320: "Purchases",
    8340: "Direct wages",
    8450: "Closing inventory",
    8518: "Cost of sales",
    8519: "Gross profit (loss)",
    8521: "Advertising and promotion",
    8570: "Amortization of intangible assets",
    8670: "Amortization of tangible assets",
    8710: "Interest and bank charges",
    8715: "Bank charges",
    8760: "Business taxes, licences and memberships",
    8810: "Office expenses",
    8860: "Professional fees",
    8861: "Legal fees",
    8862: "Accounting fees",
    8863: "Consulting fees",
    8871: "Management and administration fees",
    8910: "Rental expense",
    8960: "Repairs and maintenance",
    9060: "Salaries and wages",
    9130: "Supplies",
    9180: "Property taxes",
    9220: "Utilities",
    9225: "Telephone and telecommunications",
    9281: "Vehicle expenses",
    9284: "General and administrative expenses",
    9367: "Total operating expenses",
    9368: "Total expenses",
    9369: "Net non-farming income",
    9970: "Net income (loss) before taxes and extraordinary items",
    9990: "Current income taxes",
    9995: "Deferred income tax provision",
    9999: "Net income (loss) after taxes and extraordinary items",
}

BALANCE_TOTAL_CODES = {1599, 2008, 2009, 2178, 2179, 2599, 3139, 3450, 3499, 3620, 3640}
INCOME_TOTAL_CODES = {8299, 8518, 8519, 9367, 9368, 9369, 9970, 9999}

DEFAULT_COMPILATION_REPORT = """On the basis of information provided by management, we have compiled the balance sheet of {company_name} as at {year_end}, the statement of income for the year then ended, and Note 1, which describes the basis of accounting applied in preparing this financial information.

Management is responsible for the accompanying financial information, including the accuracy and completeness of the underlying information and the selection of the basis of accounting.

We performed this engagement in accordance with Canadian Standard on Related Services (CSRS) 4200, Compilation Engagements, and complied with relevant ethical requirements. Our responsibility is to assist management in preparing the financial information.

We did not perform an audit or review engagement and were not required to verify the accuracy or completeness of the information provided by management. Accordingly, we do not express an audit opinion, a review conclusion, or any form of assurance on this financial information.

Readers are cautioned that this financial information may not be appropriate for their purposes."""


@dataclass
class ScheduleResult:
    schedule: str
    entries: pd.DataFrame
    warnings: list[str]
    source_text: str


def _money(value: str) -> float:
    cleaned = value.strip().replace("$", "").replace(",", "").replace(" ", "")
    negative = cleaned.startswith("(") and cleaned.endswith(")")
    cleaned = cleaned.strip("()")
    if cleaned in {"", "-", "nil", "NIL"}:
        return 0.0
    number = float(cleaned)
    return -abs(number) if negative else number


def _clean_description(value: str, code: int) -> str:
    value = re.sub(r"\s+", " ", value).strip(" -:|.")
    value = re.sub(r"^(GIFI|field|line|code)\s*", "", value, flags=re.IGNORECASE)
    if not value or value.isdigit():
        return GIFI_LABELS.get(code, f"GIFI {code}")
    return value[:160]


def parse_schedule_text(text: str, schedule: str) -> ScheduleResult:
    """Extract GIFI code, description and amount rows from CRA/tax-software text."""
    schedule = schedule.upper().replace("S", "")
    if schedule not in {"100", "125"}:
        raise ValueError("Schedule must be 100 or 125.")

    amount_pattern = r"(?:\(?-?\$?\s*\d[\d,\s]*(?:\.\d{1,2})?\)?)"
    patterns = [
        re.compile(rf"^\s*(?P<code>\d{{4}})\s+(?P<desc>.*?)\s+(?P<amount>{amount_pattern})\s*$"),
        re.compile(rf"^\s*(?P<desc>.*?)\s+(?P<code>\d{{4}})\s+(?P<amount>{amount_pattern})\s*$"),
        re.compile(rf"^\s*(?P<desc>.*?)\s+(?P<amount>{amount_pattern})\s+(?P<code>\d{{4}})\s*$"),
        re.compile(rf"^\s*(?P<code>\d{{4}})\s+(?P<amount>{amount_pattern})\s*$"),
        re.compile(rf"^\s*(?P<amount>{amount_pattern})\s+(?P<code>\d{{4}})\s*$"),
    ]
    allowed = range(1000, 3650) if schedule == "100" else range(8000, 10000)
    rows: list[dict] = []
    pending_description = ""
    pending_code: int | None = None
    standalone_code_pattern = re.compile(r"^\s*(?:GIFI\s*)?(?P<code>\d{4})\s*$", re.IGNORECASE)
    standalone_amount_pattern = re.compile(
        rf"^\s*(?P<amount>{amount_pattern})(?:\s+{amount_pattern})?\s*$"
    )
    amount_token_pattern = re.compile(
        r"\(?-?\$?\s*(?:\d{1,3}(?:[,\s]\d{3})+|\d+)(?:\.\d{1,2})?\)?"
    )
    known_labels = sorted(
        (
            (code, label, re.sub(r"[^a-z0-9]+", " ", label.lower()).strip())
            for code, label in GIFI_LABELS.items()
            if code in allowed
        ),
        key=lambda item: len(item[2]),
        reverse=True,
    )

    def append_row(code: int, amount_text: str, description: str = "") -> bool:
        if code not in allowed:
            return False
        try:
            amount = _money(amount_text)
        except ValueError:
            return False
        rows.append(
            {
                "Code": code,
                "Description": _clean_description(description or pending_description, code),
                "Amount": amount,
            }
        )
        return True

    for raw_line in text.replace("\u00a0", " ").splitlines():
        line = re.sub(r"\s+", " ", raw_line).strip()
        if not line:
            continue
        match = next((pattern.match(line) for pattern in patterns if pattern.match(line)), None)
        if match:
            code = int(match.group("code"))
            if append_row(code, match.group("amount"), match.groupdict().get("desc", "")):
                pending_description = ""
                pending_code = None
            continue

        normalized_line = re.sub(r"[^a-z0-9]+", " ", line.lower()).strip()
        label_match = next(
            (
                (code, label)
                for code, label, normalized_label in known_labels
                if normalized_label and normalized_label in normalized_line
            ),
            None,
        )
        if label_match:
            code, label = label_match
            amount_candidates = amount_token_pattern.findall(line)
            # Labels do not contain numbers. If OCR also captured a GIFI code,
            # the right-most numeric token remains the statement amount.
            if amount_candidates and append_row(code, amount_candidates[-1], label):
                pending_description = ""
                pending_code = None
            else:
                pending_description = label
                pending_code = code
            continue

        code_match = standalone_code_pattern.match(line)
        if code_match:
            code = int(code_match.group("code"))
            if code in allowed:
                pending_code = code
            continue

        amount_match = standalone_amount_pattern.match(line)
        if pending_code is not None and amount_match:
            if append_row(pending_code, amount_match.group("amount")):
                pending_description = ""
                pending_code = None
            continue

        if not re.search(r"\d{4}", line) and not re.fullmatch(amount_pattern, line):
            # Positioned tax forms often extract the label, code and amount as
            # separate lines. Keep the latest text label until a row is complete.
            pending_description = line[:160]

    if not rows:
        raise RuntimeError(
            f"No Schedule {schedule} GIFI rows were found. Upload a searchable PDF exported from tax software."
        )

    entries = pd.DataFrame(rows)
    entries = entries.drop_duplicates(subset=["Code", "Amount"], keep="last")
    # If the same code appears more than once, retain the last occurrence. This
    # normally represents the final/summary statement in a tax-software PDF.
    entries = entries.drop_duplicates(subset=["Code"], keep="last").sort_values("Code").reset_index(drop=True)
    warnings = validate_schedule(entries, schedule)
    return ScheduleResult(schedule=schedule, entries=entries, warnings=warnings, source_text=text)


def extract_pdf_text(pdf_bytes: bytes) -> str:
    """Extract text using pdfplumber first, then pypdf as a fallback."""
    pages: list[str] = []
    try:
        import pdfplumber

        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages:
                pages.append(page.extract_text(x_tolerance=2, y_tolerance=3) or "")
    except Exception:
        pages = []

    text = "\n".join(pages).strip()
    if text:
        return text

    try:
        from pypdf import PdfReader

        reader = PdfReader(io.BytesIO(pdf_bytes))
        return "\n".join(page.extract_text() or "" for page in reader.pages)
    except Exception as exc:
        raise RuntimeError(f"Could not read the PDF: {exc}") from exc


def extract_pdf_layout_text(pdf_bytes: bytes) -> str:
    """Rebuild lines from positioned words when PDF text order is column-based."""
    try:
        import pdfplumber

        output: list[str] = []
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages:
                words = page.extract_words(
                    x_tolerance=2,
                    y_tolerance=3,
                    keep_blank_chars=False,
                    use_text_flow=False,
                )
                line_groups: list[list[dict]] = []
                for word in sorted(words, key=lambda item: (round(float(item["top"]) / 3), float(item["x0"]))):
                    top = float(word["top"])
                    group = next(
                        (
                            candidate
                            for candidate in line_groups
                            if abs(float(candidate[0]["top"]) - top) <= 3
                        ),
                        None,
                    )
                    if group is None:
                        group = []
                        line_groups.append(group)
                    group.append(word)
                for group in line_groups:
                    output.append(
                        " ".join(str(word["text"]) for word in sorted(group, key=lambda item: float(item["x0"])))
                    )
        return "\n".join(output)
    except Exception:
        return ""


def extract_pdf_form_text(pdf_bytes: bytes) -> str:
    """Read AcroForm and XFA values used by some T2 tax-software exports."""
    try:
        from pypdf import PdfReader

        reader = PdfReader(io.BytesIO(pdf_bytes))
        output: list[str] = []
        fields = reader.get_fields() or {}
        for name, field in fields.items():
            value = field.get("/V")
            if value in (None, ""):
                continue
            values = list(value) if isinstance(value, (list, tuple)) else [value]
            code_match = re.search(r"(?<!\d)(\d{4})(?!\d)", str(name))
            for item in values:
                if code_match:
                    output.append(f"{code_match.group(1)} {item}")
                else:
                    output.append(str(item))

        root = reader.trailer.get("/Root")
        acroform = root.get("/AcroForm") if root else None
        if acroform:
            acroform = acroform.get_object()
            xfa = acroform.get("/XFA")
            streams = []
            if isinstance(xfa, list):
                streams = [xfa[index] for index in range(1, len(xfa), 2)]
            elif xfa is not None:
                streams = [xfa]
            for stream in streams:
                try:
                    xml = stream.get_object().get_data().decode("utf-8", errors="ignore")
                    output.append(re.sub(r"<[^>]+>", "\n", xml))
                except Exception:
                    continue
        return "\n".join(output)
    except Exception:
        return ""


def extract_docling_text(pdf_bytes: bytes) -> str:
    """Use Docling as the final OCR/table fallback for graphical tax PDFs."""
    try:
        from docling.document_converter import DocumentConverter

        with tempfile.TemporaryDirectory(prefix="gifi_docling_") as temp_name:
            pdf_path = Path(temp_name) / "schedule.pdf"
            pdf_path.write_bytes(pdf_bytes)
            document = DocumentConverter().convert(str(pdf_path)).document
            for method_name in ("export_to_markdown", "export_to_text"):
                method = getattr(document, method_name, None)
                if method:
                    text = method()
                    if text and str(text).strip():
                        return str(text)
    except Exception:
        return ""
    return ""


def extract_tesseract_text(pdf_bytes: bytes) -> str:
    """Render pages at 300 DPI and OCR PDFs that contain no usable text layer."""
    try:
        import fitz
        import pytesseract
        from PIL import Image, ImageOps

        document = fitz.open(stream=pdf_bytes, filetype="pdf")
        output: list[str] = []
        scale = 300 / 72
        matrix = fitz.Matrix(scale, scale)
        for page in document:
            pixmap = page.get_pixmap(matrix=matrix, alpha=False, colorspace=fitz.csRGB)
            image = Image.open(io.BytesIO(pixmap.tobytes("png")))
            grayscale = ImageOps.autocontrast(ImageOps.grayscale(image))
            output.append(pytesseract.image_to_string(grayscale, config="--oem 3 --psm 6"))
            thresholded = grayscale.point(lambda value: 0 if value < 185 else 255)
            output.append(pytesseract.image_to_string(thresholded, config="--oem 3 --psm 11"))
        document.close()
        return "\n".join(output)
    except Exception:
        return ""


def extract_schedule_pdf(pdf_bytes: bytes, schedule: str) -> ScheduleResult:
    extracted_sources: list[str] = []
    for extractor in (extract_pdf_text, extract_pdf_layout_text, extract_pdf_form_text):
        try:
            text = extractor(pdf_bytes)
        except RuntimeError:
            text = ""
        if not text or text in extracted_sources:
            continue
        extracted_sources.append(text)
        try:
            return parse_schedule_text(text, schedule)
        except RuntimeError:
            continue

    tesseract_text = extract_tesseract_text(pdf_bytes)
    if tesseract_text:
        try:
            return parse_schedule_text(tesseract_text, schedule)
        except RuntimeError:
            extracted_sources.append(tesseract_text)

    docling_text = extract_docling_text(pdf_bytes)
    if docling_text:
        try:
            return parse_schedule_text(docling_text, schedule)
        except RuntimeError:
            extracted_sources.append(docling_text)

    raise RuntimeError(
        f"No Schedule {str(schedule).upper().replace('S', '')} GIFI rows were found after "
        "text, form-field, positioned-layout, Tesseract OCR, and Docling extraction. "
        "Upload the original PDF exported from the tax software, not a print preview."
    )


def _value(entries: pd.DataFrame, code: int) -> float | None:
    match = entries.loc[entries["Code"] == code, "Amount"]
    return None if match.empty else float(match.iloc[-1])


def validate_schedule(entries: pd.DataFrame, schedule: str) -> list[str]:
    warnings: list[str] = []
    if schedule == "100":
        assets = _value(entries, 2599)
        liabilities = _value(entries, 3499)
        equity = _value(entries, 3620)
        combined = _value(entries, 3640)
        if assets is None:
            warnings.append("Schedule 100 is missing GIFI 2599 Total assets.")
        if liabilities is None:
            warnings.append("Schedule 100 is missing GIFI 3499 Total liabilities.")
        if equity is None:
            warnings.append("Schedule 100 is missing GIFI 3620 Total shareholder equity.")
        if assets is not None and liabilities is not None and equity is not None:
            if abs(assets - (liabilities + equity)) > 1:
                warnings.append(
                    f"Balance sheet does not balance: assets ${assets:,.2f} versus "
                    f"liabilities and equity ${liabilities + equity:,.2f}."
                )
        if assets is not None and combined is not None and abs(assets - combined) > 1:
            warnings.append("GIFI 2599 does not agree with GIFI 3640.")
    else:
        revenue = _value(entries, 8299)
        expenses = _value(entries, 9368)
        income = _value(entries, 9970)
        if revenue is None:
            warnings.append("Schedule 125 is missing GIFI 8299 Total revenue.")
        if expenses is None:
            warnings.append("Schedule 125 is missing GIFI 9368 Total expenses.")
        if income is None:
            warnings.append("Schedule 125 is missing GIFI 9970 Net income before taxes.")
        if revenue is not None and expenses is not None and income is not None:
            if abs((revenue - expenses) - income) > 1:
                warnings.append(
                    f"Income statement does not reconcile: revenue less expenses is "
                    f"${revenue - expenses:,.2f}, but GIFI 9970 is ${income:,.2f}."
                )
    return warnings


def statement_tables(balance: pd.DataFrame, income: pd.DataFrame) -> dict[str, pd.DataFrame]:
    assets = balance[(balance["Code"] < 2600) & ~balance["Code"].isin(BALANCE_TOTAL_CODES)].copy()
    liabilities = balance[
        (balance["Code"] >= 2600) & (balance["Code"] < 3500) & ~balance["Code"].isin(BALANCE_TOTAL_CODES)
    ].copy()
    equity = balance[
        (balance["Code"] >= 3500) & (balance["Code"] < 3650) & ~balance["Code"].isin(BALANCE_TOTAL_CODES)
    ].copy()
    revenue = income[(income["Code"] < 8300) & ~income["Code"].isin(INCOME_TOTAL_CODES)].copy()
    cost_of_sales = income[
        (income["Code"] >= 8300) & (income["Code"] < 8518) & ~income["Code"].isin(INCOME_TOTAL_CODES)
    ].copy()
    expenses = income[
        (income["Code"] >= 8520) & (income["Code"] < 9970) & ~income["Code"].isin(INCOME_TOTAL_CODES)
    ].copy()
    taxes = income[(income["Code"] >= 9975) & (income["Code"] < 9999)].copy()
    return {
        "assets": assets,
        "liabilities": liabilities,
        "equity": equity,
        "revenue": revenue,
        "cost_of_sales": cost_of_sales,
        "expenses": expenses,
        "taxes": taxes,
    }


def _format_amount(value: float) -> str:
    return f"({abs(value):,.0f})" if value < 0 else f"{value:,.0f}"


def _report_text(metadata: dict) -> str:
    template = metadata.get("report_text", "").strip() or DEFAULT_COMPILATION_REPORT
    return template.replace(
        "{company_name}", metadata["company_name"]
    ).replace(
        "{year_end}", metadata["year_end"].strftime("%B %d, %Y")
    )


def build_financial_statement_pdf(balance: pd.DataFrame, income: pd.DataFrame, metadata: dict) -> bytes:
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER, TA_RIGHT
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import inch
    from reportlab.platypus import (
        KeepTogether,
        PageBreak,
        Paragraph,
        SimpleDocTemplate,
        Spacer,
        Table,
        TableStyle,
    )

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        rightMargin=0.75 * inch,
        leftMargin=0.75 * inch,
        topMargin=0.65 * inch,
        bottomMargin=0.65 * inch,
        title=f"{metadata['company_name']} Compiled Financial Information",
    )
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="Company", parent=styles["Title"], fontName="Helvetica-Bold", fontSize=16, leading=19, alignment=TA_CENTER, spaceAfter=5))
    styles.add(ParagraphStyle(name="StatementTitle", parent=styles["Heading1"], fontName="Helvetica-Bold", fontSize=12, leading=15, alignment=TA_CENTER, spaceAfter=3))
    styles.add(ParagraphStyle(name="Period", parent=styles["Normal"], fontSize=9.5, alignment=TA_CENTER, spaceAfter=14))
    styles.add(ParagraphStyle(name="Draft", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=8, textColor=colors.HexColor("#A3262A"), alignment=TA_CENTER, spaceAfter=10))
    styles.add(ParagraphStyle(name="ReportBody", parent=styles["Normal"], fontSize=9.5, leading=14, spaceAfter=10))
    styles.add(ParagraphStyle(name="NoteBody", parent=styles["Normal"], fontSize=9.5, leading=14, spaceAfter=8))
    styles.add(ParagraphStyle(name="TableText", parent=styles["Normal"], fontName="Helvetica", fontSize=9, leading=11))

    year_end = metadata["year_end"].strftime("%B %d, %Y")
    tables = statement_tables(balance, income)

    def page_header(title: str, period: str) -> list:
        result = [
            Paragraph(escape(metadata["company_name"]), styles["Company"]),
            Paragraph(escape(title), styles["StatementTitle"]),
            Paragraph(escape(period), styles["Period"]),
        ]
        if metadata.get("business_number"):
            result.append(
                Paragraph(f"Business number: {escape(metadata['business_number'])}", styles["Period"])
            )
        return result

    def financial_table(sections: Iterable[tuple[str, pd.DataFrame, tuple[str, float] | None]]) -> Table:
        rows: list[list] = [["", year_end]]
        bold_rows = {0}
        for heading, frame, total in sections:
            rows.append([heading, ""])
            bold_rows.add(len(rows) - 1)
            for row in frame.itertuples(index=False):
                rows.append([row.Description, _format_amount(float(row.Amount))])
            if total:
                rows.append([total[0], _format_amount(float(total[1]))])
                bold_rows.add(len(rows) - 1)
                rows.append(["", ""])
        display_rows = []
        for row_index, row in enumerate(rows):
            label = escape(str(row[0]))
            if row_index in bold_rows:
                label = f"<b>{label}</b>"
            display_rows.append([Paragraph(label, styles["TableText"]), row[1]])
        table = Table(display_rows, colWidths=[5.15 * inch, 1.35 * inch], repeatRows=1)
        commands = [
            ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
            ("FONTSIZE", (0, 0), (-1, -1), 9),
            ("ALIGN", (1, 0), (1, -1), "RIGHT"),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("LINEBELOW", (1, 0), (1, 0), 0.75, colors.black),
        ]
        for row_index in bold_rows:
            commands.append(("FONTNAME", (0, row_index), (-1, row_index), "Helvetica-Bold"))
        table.setStyle(TableStyle(commands))
        return table

    story = []
    story.extend(page_header("Compilation Engagement Report", ""))
    story.append(Paragraph("DRAFT - FOR PRACTITIONER REVIEW AND SIGNATURE", styles["Draft"]))
    story.append(
        Paragraph(f"To Management of {escape(metadata['company_name'])}", styles["ReportBody"])
    )
    for paragraph in _report_text(metadata).split("\n\n"):
        story.append(Paragraph(escape(paragraph), styles["ReportBody"]))
    story.extend(
        [
            Spacer(1, 16),
            Paragraph(escape(metadata["firm_name"]), styles["ReportBody"]),
            Paragraph(escape(metadata["firm_address"]), styles["ReportBody"]),
            Paragraph(metadata["report_date"].strftime("%B %d, %Y"), styles["ReportBody"]),
            PageBreak(),
        ]
    )

    story.extend(page_header("Balance Sheet", f"As at {year_end}"))
    story.append(
        financial_table(
            [
                ("ASSETS", tables["assets"], ("Total assets", _value(balance, 2599) or 0)),
                ("LIABILITIES", tables["liabilities"], ("Total liabilities", _value(balance, 3499) or 0)),
                ("SHAREHOLDER EQUITY", tables["equity"], ("Total shareholder equity", _value(balance, 3620) or 0)),
                (
                    "",
                    pd.DataFrame(columns=balance.columns),
                    ("Total liabilities and shareholder equity", _value(balance, 3640) or ((_value(balance, 3499) or 0) + (_value(balance, 3620) or 0))),
                ),
            ]
        )
    )
    story.append(PageBreak())

    story.extend(page_header("Statement of Income", f"For the year ended {year_end}"))
    income_sections = [
        ("REVENUE", tables["revenue"], ("Total revenue", _value(income, 8299) or 0)),
    ]
    if not tables["cost_of_sales"].empty or _value(income, 8518) is not None:
        income_sections.append(
            ("COST OF SALES", tables["cost_of_sales"], ("Cost of sales", _value(income, 8518) or 0))
        )
        gross_profit = _value(income, 8519)
        if gross_profit is None:
            gross_profit = (_value(income, 8299) or 0) - (_value(income, 8518) or 0)
        income_sections.append(("", pd.DataFrame(columns=income.columns), ("Gross profit (loss)", gross_profit)))
    income_sections.append(
        ("OPERATING EXPENSES", tables["expenses"], ("Total operating expenses", _value(income, 9367) or 0))
    )
    income_sections.append(
        ("", pd.DataFrame(columns=income.columns), ("Total expenses", _value(income, 9368) or 0))
    )
    if not tables["taxes"].empty:
        income_sections.append(("INCOME TAXES AND OTHER ITEMS", tables["taxes"], None))
    income_sections.append(("", pd.DataFrame(columns=income.columns), ("Net income (loss)", _value(income, 9999) if _value(income, 9999) is not None else (_value(income, 9970) or 0))))
    story.append(financial_table(income_sections))
    story.append(PageBreak())

    story.extend(page_header("Notes to Compiled Financial Information", f"For the year ended {year_end}"))
    story.append(
        KeepTogether(
            [
                Paragraph("1. Basis of accounting", styles["Heading2"]),
                Paragraph(escape(metadata["basis_of_accounting"]), styles["NoteBody"]),
            ]
        )
    )
    story.append(Spacer(1, 12))
    story.append(Paragraph("The accompanying financial information was compiled from Schedule 100 and Schedule 125 GIFI information supplied by management. The practitioner must review classifications, disclosures, and the report wording before issuance.", styles["NoteBody"]))

    def footer(canvas, _doc):
        canvas.saveState()
        if _doc.page > 1:
            canvas.setFillColor(colors.HexColor("#A3262A"))
            canvas.setFont("Helvetica-Bold", 8)
            canvas.drawCentredString(
                letter[0] / 2,
                0.5 * inch,
                "Unaudited - See Compilation Engagement Report",
            )
            canvas.setFillColor(colors.black)
        canvas.setFont("Helvetica", 8)
        canvas.drawCentredString(letter[0] / 2, 0.35 * inch, str(_doc.page))
        canvas.restoreState()

    doc.build(story, onFirstPage=footer, onLaterPages=footer)
    return buffer.getvalue()


def build_financial_statement_docx(balance: pd.DataFrame, income: pd.DataFrame, metadata: dict) -> bytes:
    from docx import Document
    from docx.enum.section import WD_SECTION
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Inches, Pt, RGBColor

    document = Document()
    section = document.sections[0]
    section.top_margin = Inches(0.7)
    section.bottom_margin = Inches(0.7)
    section.left_margin = Inches(0.8)
    section.right_margin = Inches(0.8)
    normal = document.styles["Normal"]
    normal.font.name = "Arial"
    normal.font.size = Pt(10)

    year_end = metadata["year_end"].strftime("%B %d, %Y")
    tables = statement_tables(balance, income)

    def add_heading_block(title: str, period: str = ""):
        p = document.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(metadata["company_name"])
        run.bold = True
        run.font.name = "Arial"
        run.font.size = Pt(16)
        p = document.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(title)
        run.bold = True
        run.font.size = Pt(12)
        if period:
            p = document.add_paragraph(period)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if metadata.get("business_number"):
            p = document.add_paragraph(f"Business number: {metadata['business_number']}")
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    def add_statement_table(sections):
        rows = [["", year_end]]
        bold = {0}
        for heading, frame, total in sections:
            rows.append([heading, ""])
            bold.add(len(rows) - 1)
            for item in frame.itertuples(index=False):
                rows.append([item.Description, _format_amount(float(item.Amount))])
            if total:
                rows.append([total[0], _format_amount(float(total[1]))])
                bold.add(len(rows) - 1)
                rows.append(["", ""])
        table = document.add_table(rows=0, cols=2)
        table.autofit = False
        for index, row in enumerate(rows):
            cells = table.add_row().cells
            cells[0].width = Inches(5.2)
            cells[1].width = Inches(1.3)
            cells[0].text = row[0]
            cells[1].text = row[1]
            cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
            if index in bold:
                for cell in cells:
                    for run in cell.paragraphs[0].runs:
                        run.bold = True
        table.style = "Table Grid"

    add_heading_block("Compilation Engagement Report")
    p = document.add_paragraph("DRAFT - FOR PRACTITIONER REVIEW AND SIGNATURE")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in p.runs:
        run.bold = True
        run.font.color.rgb = RGBColor(163, 38, 42)
        run.font.size = Pt(8)
    document.add_paragraph(f"To Management of {metadata['company_name']}")
    for paragraph in _report_text(metadata).split("\n\n"):
        document.add_paragraph(paragraph)
    document.add_paragraph(metadata["firm_name"])
    document.add_paragraph(metadata["firm_address"])
    document.add_paragraph(metadata["report_date"].strftime("%B %d, %Y"))
    document.add_page_break()

    add_heading_block("Balance Sheet", f"As at {year_end}")
    add_statement_table(
        [
            ("ASSETS", tables["assets"], ("Total assets", _value(balance, 2599) or 0)),
            ("LIABILITIES", tables["liabilities"], ("Total liabilities", _value(balance, 3499) or 0)),
            ("SHAREHOLDER EQUITY", tables["equity"], ("Total shareholder equity", _value(balance, 3620) or 0)),
            ("", pd.DataFrame(columns=balance.columns), ("Total liabilities and shareholder equity", _value(balance, 3640) or 0)),
        ]
    )
    document.add_paragraph("Unaudited - See Compilation Engagement Report").alignment = WD_ALIGN_PARAGRAPH.CENTER
    document.add_page_break()

    add_heading_block("Statement of Income", f"For the year ended {year_end}")
    sections = [("REVENUE", tables["revenue"], ("Total revenue", _value(income, 8299) or 0))]
    if not tables["cost_of_sales"].empty or _value(income, 8518) is not None:
        sections.append(("COST OF SALES", tables["cost_of_sales"], ("Cost of sales", _value(income, 8518) or 0)))
        sections.append(("", pd.DataFrame(columns=income.columns), ("Gross profit (loss)", _value(income, 8519) or 0)))
    sections.append(("OPERATING EXPENSES", tables["expenses"], ("Total operating expenses", _value(income, 9367) or 0)))
    sections.append(("", pd.DataFrame(columns=income.columns), ("Total expenses", _value(income, 9368) or 0)))
    if not tables["taxes"].empty:
        sections.append(("INCOME TAXES AND OTHER ITEMS", tables["taxes"], None))
    sections.append(("", pd.DataFrame(columns=income.columns), ("Net income (loss)", _value(income, 9999) if _value(income, 9999) is not None else (_value(income, 9970) or 0))))
    add_statement_table(sections)
    document.add_paragraph("Unaudited - See Compilation Engagement Report").alignment = WD_ALIGN_PARAGRAPH.CENTER
    document.add_page_break()

    add_heading_block("Notes to Compiled Financial Information", f"For the year ended {year_end}")
    document.add_heading("1. Basis of accounting", level=2)
    document.add_paragraph(metadata["basis_of_accounting"])
    document.add_paragraph(
        "The accompanying financial information was compiled from Schedule 100 and Schedule 125 "
        "GIFI information supplied by management. The practitioner must review classifications, "
        "disclosures, and report wording before issuance."
    )

    for section in document.sections:
        footer = section.footer.paragraphs[0]
        footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
        footer.add_run("Draft compiled financial information")

    output = io.BytesIO()
    document.save(output)
    return output.getvalue()


def suggested_filename(company_name: str, year_end: date, extension: str) -> str:
    safe = re.sub(r"[^A-Za-z0-9_-]+", "_", company_name).strip("_") or "Company"
    return f"{safe}_{year_end.isoformat()}_Compiled_Financial_Statements.{extension}"
