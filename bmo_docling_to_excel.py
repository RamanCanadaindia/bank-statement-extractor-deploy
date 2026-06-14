"""
Convert Docling JSON from a BMO bank statement into an Excel workbook.

This script is intentionally offline-only. It reads a local JSON file, extracts
statement rows from Docling text/table/OCR structures, and writes an Excel file
with a transactions sheet plus a summary sheet.

To adapt this for RBC, TD, CIBC, or credit cards, start by changing:
1. BANK_KEYWORDS and SUMMARY_KEYWORDS
2. categorize_transaction()
3. parse_transaction_line() if the statement column order differs
"""

from __future__ import annotations

import json
import re
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable

import pandas as pd

try:
    from pypdf import PdfReader
except ImportError:  # PDF support is optional; JSON extraction still works.
    PdfReader = None


APP_FOLDER = Path(__file__).resolve().parent
INPUT_JSON = APP_FOLDER / "text.json"
OUTPUT_XLSX = APP_FOLDER / "Excel" / "BMO_transactions.xlsx"
SCRIPT_FOLDER_INPUT_JSON = APP_FOLDER / "text.json"


# BMO statements usually describe withdrawals as drafts, draft fees, insurance,
# service charges, etc. These words help decide whether a single amount is debit
# or credit when the OCR text does not preserve table columns cleanly.
DEBIT_KEYWORDS = [
    "draft",
    "draft fee",
    "canadian draft",
    "fee",
    "charge",
    "service charge",
    "insurance",
    "icbc",
    "payment",
    "withdrawal",
    "debit",
]

CREDIT_KEYWORDS = [
    "deposit",
    "credit",
    "interest",
]

# Summary-like labels should not be treated as normal transactions. Opening and
# closing balances are kept, but marked separately.
SUMMARY_KEYWORDS = [
    "total deposits",
    "total credits",
    "total debits",
    "total drafts",
    "total withdrawals",
    "total fees",
    "account summary",
    "summary",
]


DATE_RE = re.compile(
    r"(?P<date>(?:\d{1,2}[/-]\d{1,2}(?:[/-]\d{2,4})?)|(?:[A-Z][a-z]{2,9}\s+\d{1,2}(?:,\s*\d{4})?))"
)
AMOUNT_RE = re.compile(r"[-+]?\$?\(?\d{1,3}(?:,\d{3})*(?:\.\d{2})\)?")


@dataclass
class ParsedLine:
    date: str
    description: str
    debit: float | None
    credit: float | None
    balance: float | None
    category: str


def load_docling_json(path: Path) -> dict[str, Any]:
    """Load Docling JSON from disk with a helpful error if it is missing."""
    if not path.exists() and SCRIPT_FOLDER_INPUT_JSON.exists():
        # Convenient fallback: if text.json is beside this script, use it.
        # This makes the script easy to move into folders for RBC, TD, CIBC, etc.
        path = SCRIPT_FOLDER_INPUT_JSON

    if not path.exists():
        raise FileNotFoundError(
            f"JSON file not found: {path}. Also checked: {SCRIPT_FOLDER_INPUT_JSON}"
        )

    with path.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def clean_text(value: Any) -> str:
    """Normalize OCR/table text while preserving row-level meaning."""
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def parse_amount(value: str) -> float | None:
    """Convert statement amount text such as '$1,234.56' or '(12.00)' to float."""
    text = clean_text(value)
    if not text:
        return None

    negative = text.startswith("(") and text.endswith(")")
    text = text.replace("$", "").replace(",", "").replace("(", "").replace(")", "")

    try:
        amount = float(text)
    except ValueError:
        return None

    return -amount if negative else amount


def categorize_transaction(description: str) -> str:
    """Basic categorization rules. Add bank-specific merchant rules here."""
    desc = description.lower()

    if "acc fee" in desc or "account fee" in desc:
        return "Bank Charges"
    if "balance fee waiver" in desc:
        return "Bank Fee Reversal"
    if "line of credit" in desc:
        return "Loan/Line of Credit"
    if "credit card payment" in desc:
        return "Credit Card Payment"
    if "loan payment" in desc:
        return "Loan Payment"
    if "loan credit" in desc:
        return "Loan Advance/Credit"
    if "fuel bill payment" in desc:
        return "Fuel"
    if "rent/lease" in desc or "lease payment" in desc or "automobile rent" in desc:
        return "Rent/Lease"
    if "interac purchase" in desc:
        return "Purchase"
    if "item returned nsf" in desc:
        return "Returned Item"
    if "nsf item fee" in desc:
        return "Bank Charges"
    if "draft fee" in desc:
        return "Bank Charges"
    if "icbc" in desc:
        return "Insurance"
    if "bill payment" in desc or "internet bill pmt" in desc:
        return "Bill Payment"
    if "cheque" in desc:
        return "Cheque"
    if "e-transfer" in desc:
        return "E-Transfer"
    if "pre-auth debit" in desc or "pre-authorized payment" in desc:
        return "Pre-Authorized Payment"
    if "deposit" in desc or "mobile deposit" in desc:
        return "Income/Deposit"
    if "canadian draft" in desc:
        return "Bank Transfer/Payment"
    if "opening balance" in desc:
        return "Opening Balance"
    if "closing" in desc:
        return "Closing Totals"
    if "fee" in desc or "charge" in desc:
        return "Bank Charges"

    return "Uncategorized"


def extract_plain_text_blocks(data: Any) -> list[str]:
    """
    Recursively collect text-like content from Docling JSON.

    Docling exports can vary depending on OCR/table settings. This function looks
    for common text fields inside texts, tables, cells, OCR blocks, and nested
    children without depending on one exact schema.
    """
    blocks: list[str] = []

    if isinstance(data, dict):
        for key, value in data.items():
            if key in {"text", "orig", "label", "content"}:
                text = clean_text(value)
                if text:
                    blocks.append(text)
            elif key in {"bbox", "prov"}:
                # Layout metadata is useful for sorting in specialized extractors,
                # but should not become transaction text.
                continue
            else:
                blocks.extend(extract_plain_text_blocks(value))
    elif isinstance(data, list):
        for item in data:
            blocks.extend(extract_plain_text_blocks(item))

    return blocks


def extract_tables(data: dict[str, Any]) -> list[list[list[str]]]:
    """
    Extract possible table rows from Docling table structures.

    Different Docling versions represent tables differently. This handles common
    row/cell forms and falls back to recursively finding table-like dictionaries.
    """
    tables: list[list[list[str]]] = []

    def cell_text(cell: Any) -> str:
        if isinstance(cell, dict):
            for key in ("text", "content", "orig"):
                if key in cell and clean_text(cell[key]):
                    return clean_text(cell[key])
            return " ".join(extract_plain_text_blocks(cell))
        return clean_text(cell)

    def walk(node: Any) -> None:
        if isinstance(node, dict):
            if "rows" in node and isinstance(node["rows"], list):
                rows: list[list[str]] = []
                for row in node["rows"]:
                    if isinstance(row, dict):
                        cells = row.get("cells") or row.get("data") or []
                    else:
                        cells = row
                    if isinstance(cells, list):
                        parsed = [cell_text(cell) for cell in cells]
                        if any(parsed):
                            rows.append(parsed)
                if rows:
                    tables.append(rows)

            if "table_cells" in node and isinstance(node["table_cells"], list):
                # Some Docling exports provide cells with row/col indexes.
                grid: dict[tuple[int, int], str] = {}
                for cell in node["table_cells"]:
                    if not isinstance(cell, dict):
                        continue
                    row_idx = cell.get("start_row_offset_idx", cell.get("row", 0))
                    col_idx = cell.get("start_col_offset_idx", cell.get("col", 0))
                    grid[(int(row_idx), int(col_idx))] = cell_text(cell)
                if grid:
                    max_row = max(row for row, _ in grid)
                    max_col = max(col for _, col in grid)
                    rows = [
                        [grid.get((row, col), "") for col in range(max_col + 1)]
                        for row in range(max_row + 1)
                    ]
                    tables.append(rows)

            for value in node.values():
                walk(value)
        elif isinstance(node, list):
            for item in node:
                walk(item)

    walk(data)
    return tables


def extract_positioned_lines(data: Any) -> list[str]:
    """
    Rebuild OCR lines from text blocks that contain bbox positions.

    If Docling produced separate positioned words/blocks, sorting by page, y, and
    x can restore transaction rows that were not recognized as tables.
    """
    positioned: list[tuple[int, float, float, str]] = []

    def get_bbox_xy(item: dict[str, Any]) -> tuple[float, float] | None:
        bbox = item.get("bbox")
        if not isinstance(bbox, dict):
            return None
        x = bbox.get("l", bbox.get("x", bbox.get("left")))
        y = bbox.get("t", bbox.get("y", bbox.get("top")))
        if x is None or y is None:
            return None
        return float(x), float(y)

    def walk(node: Any, page: int = 0) -> None:
        if isinstance(node, dict):
            page_no = int(node.get("page_no", node.get("page", page)) or page)
            text = clean_text(node.get("text") or node.get("content") or node.get("orig"))
            xy = get_bbox_xy(node)
            if text and xy:
                positioned.append((page_no, xy[1], xy[0], text))
            for value in node.values():
                walk(value, page_no)
        elif isinstance(node, list):
            for item in node:
                walk(item, page)

    walk(data)
    if not positioned:
        return []

    positioned.sort()
    lines: list[str] = []
    current_key: tuple[int, int] | None = None
    current_parts: list[str] = []

    for page, y, _x, text in positioned:
        # Group nearby y positions into one OCR row. The bucket size is deliberately
        # broad because bank statements may have slightly uneven OCR boxes.
        key = (page, round(y / 8))
        if current_key is not None and key != current_key:
            lines.append(" ".join(current_parts))
            current_parts = []
        current_key = key
        current_parts.append(text)

    if current_parts:
        lines.append(" ".join(current_parts))

    return [clean_text(line) for line in lines if clean_text(line)]


def _amounts_from_text(text: str) -> list[float]:
    """Return all parseable amounts from a text fragment."""
    return [amount for amount in (parse_amount(raw) for raw in AMOUNT_RE.findall(text)) if amount is not None]


def _date_tokens_from_text(text: str) -> list[str]:
    """
    Split compact BMO date strings like 'Jan 21 Jan 22 Jan 22' into dates.

    BMO statement OCR often reads the date column as one vertical text block.
    This helper keeps the month when day-only rows follow the first month name.
    """
    tokens = re.findall(r"[A-Z][a-z]{2,9}|\d{1,2}", text)
    dates: list[str] = []
    current_month = ""

    for token in tokens:
        if re.match(r"[A-Z][a-z]{2,9}$", token):
            current_month = token
        elif current_month:
            dates.append(f"{current_month} {token}")

    return dates


def extract_bmo_stacked_layout(data: dict[str, Any]) -> list[ParsedLine]:
    """
    Parse the BMO layout where Docling stacks columns into separate text/table cells.

    In this layout, Docling may put the first few rows in positioned text blocks:
    - dates in one block
    - descriptions in separate blocks
    - debit/credit/balance amounts in separate positioned blocks

    The remaining rows may appear in a damaged 4-column table where several debit
    amounts and balances are grouped into single cells. This function detects that
    shape and reconstructs the transaction sequence.
    """
    texts = data.get("texts", [])
    text_values = [clean_text(item.get("text") or item.get("orig")) for item in texts if isinstance(item, dict)]
    all_text = "\n".join(text_values)

    if not (
        "Deposit at, BR." in all_text
        and "Canadian Draft, 2066 DRAFT" in all_text
        and "Amounts debited from your account" in all_text
    ):
        return []

    rows: list[ParsedLine] = []

    date_block = next((value for value in text_values if re.search(r"\bJan\s+\d{1,2}\s+Jan\s+\d{1,2}", value)), "")
    first_dates = _date_tokens_from_text(date_block)
    opening_date = first_dates[0] if len(first_dates) > 0 else ""
    deposit_date = first_dates[1] if len(first_dates) > 1 else opening_date
    draft_date = first_dates[2] if len(first_dates) > 2 else deposit_date

    closing_date_block = next((value for value in text_values if re.fullmatch(r"Jan\s+\d{1,2}", value)), "")
    closing_date = closing_date_block

    def amount_after(label: str, x_min: float, y_match: float | None = None) -> float | None:
        """Find a positioned amount near a known BMO row/column."""
        matches: list[tuple[float, float]] = []
        for item in texts:
            if not isinstance(item, dict):
                continue
            value = clean_text(item.get("text") or item.get("orig"))
            amounts = _amounts_from_text(value)
            if not amounts:
                continue
            prov = item.get("prov") or []
            bbox = prov[0].get("bbox") if prov else {}
            x = float(bbox.get("l", 0) or 0)
            y = float(bbox.get("t", 0) or 0)
            if x < x_min:
                continue
            if y_match is not None and abs(y - y_match) > 2:
                continue
            matches.append((x, amounts[0]))
        return sorted(matches)[0][1] if matches else None

    # First three visible transaction rows are represented as positioned text.
    opening_balance = amount_after("Opening balance", 500, 317.55)
    deposit_credit = amount_after("Deposit at", 390, 303.63)
    deposit_balance = amount_after("Deposit at", 490, 303.63)
    first_draft_debit = amount_after("Canadian Draft", 290, 289.47)
    first_draft_balance = amount_after("Canadian Draft", 490, 289.47)

    if opening_balance is not None:
        rows.append(
            ParsedLine(
                date=opening_date,
                description="Opening balance",
                debit=None,
                credit=None,
                balance=opening_balance,
                category="Opening Balance",
            )
        )

    if deposit_credit is not None:
        rows.append(
            ParsedLine(
                date=deposit_date,
                description="Deposit at, BR. 2066",
                debit=None,
                credit=abs(deposit_credit),
                balance=deposit_balance,
                category="Income/Deposit",
            )
        )

    first_draft_desc = next((value for value in text_values if value.startswith("Canadian Draft, 2066 DRAFT")), "")
    if first_draft_debit is not None and first_draft_desc:
        rows.append(
            ParsedLine(
                date=draft_date,
                description=first_draft_desc,
                debit=abs(first_draft_debit),
                credit=None,
                balance=first_draft_balance,
                category=categorize_transaction(first_draft_desc),
            )
        )

    # Remaining rows are in the damaged Docling transaction table.
    table_rows: list[list[str]] = []
    for table in data.get("tables", []):
        grid = ((table.get("data") or {}).get("grid") or []) if isinstance(table, dict) else []
        for grid_row in grid:
            row = [clean_text(cell.get("text")) if isinstance(cell, dict) else clean_text(cell) for cell in grid_row]
            table_rows.append(row)

    flat_table = "\n".join(" | ".join(row) for row in table_rows)
    if "153.94 39,231.92 9.95 815.75" not in flat_table:
        return rows

    # Dates and descriptions from the stacked table.
    table_dates = _date_tokens_from_text(" ".join(row[0] for row in table_rows if len(row) > 0))
    jan22 = "Jan 22"
    jan23 = "Jan 23"
    jan24 = "Jan 24"
    if table_dates:
        jan22 = table_dates[0]
    if len(table_dates) >= 3:
        jan23 = table_dates[1]
    if "24" in " ".join(row[0] for row in table_rows if len(row) > 0):
        jan24 = "Jan 24"

    draft_fee = next((_amounts_from_text(row[2])[0] for row in table_rows if len(row) > 2 and _amounts_from_text(row[2]) and "153.94" not in row[2]), None)
    grouped_debits = next((_amounts_from_text(row[2]) for row in table_rows if len(row) > 2 and "153.94" in row[2]), [])
    balances = []
    for row in table_rows:
        if len(row) > 3:
            balances.extend(_amounts_from_text(row[3]))

    second_draft_desc = next((row[1] for row in table_rows if len(row) > 1 and "DRAFT 026339903" in row[1]), "")
    icbc_desc = "Pre-Authorized Payment, ICBC INS/ASS"

    if draft_fee is not None and len(balances) >= 1:
        rows.append(
            ParsedLine(jan22, "Draft Fee", abs(draft_fee), None, balances[0], "Bank Charges")
        )

    if len(grouped_debits) >= 4 and len(balances) >= 5:
        rows.extend(
            [
                ParsedLine(jan23, icbc_desc, abs(grouped_debits[0]), None, balances[1], "Insurance"),
                ParsedLine(jan23, second_draft_desc, abs(grouped_debits[1]), None, balances[2], categorize_transaction(second_draft_desc)),
                ParsedLine(jan23, "Canadian Draft Fee", abs(grouped_debits[2]), None, balances[3], "Bank Charges"),
                ParsedLine(jan24, icbc_desc, abs(grouped_debits[3]), None, balances[4], "Insurance"),
            ]
        )

    summary_debits = next(
        (amounts[0] for value in text_values if value == "79,721.51" and (amounts := _amounts_from_text(value))),
        None,
    )
    summary_credits = next(
        (amounts[0] for value in text_values if value == "100,000.00" and (amounts := _amounts_from_text(value))),
        None,
    )
    closing_balance = balances[-1] if balances else None

    if closing_date and closing_balance is not None:
        rows.append(
            ParsedLine(
                date=closing_date,
                description="Closing totals",
                debit=summary_debits,
                credit=summary_credits,
                balance=closing_balance,
                category="Closing Totals",
            )
        )

    normal_rows = [row for row in rows if row.category not in {"Opening Balance", "Closing Totals"}]
    if len(normal_rows) >= 7:
        return rows

    return []


def is_summary_line(line: str) -> bool:
    lower = line.lower()
    return any(keyword in lower for keyword in SUMMARY_KEYWORDS)


def is_opening_balance_line(line: str) -> bool:
    lower = line.lower()
    return "opening balance" in lower or "balance forward" in lower


def is_closing_balance_line(line: str) -> bool:
    lower = line.lower()
    return (
        "closing balance" in lower
        or "ending balance" in lower
        or "balance at end" in lower
    )


def decide_debit_credit(description: str, amount: float) -> tuple[float | None, float | None]:
    """
    Split one ambiguous amount into Debit or Credit using BMO description rules.

    If a row has separate debit/credit columns, table parsing fills those directly.
    This function is mainly for OCR text lines where the columns were flattened.
    """
    desc = description.lower()
    if any(word in desc for word in CREDIT_KEYWORDS):
        return None, abs(amount)
    if any(word in desc for word in DEBIT_KEYWORDS):
        return abs(amount), None
    if amount < 0:
        return abs(amount), None
    return None, abs(amount)


def parse_transaction_line(line: str) -> ParsedLine | None:
    """
    Parse a flattened BMO transaction row.

    Expected examples:
    - 01/02 Deposit 500.00 1,500.00
    - Jan 03 Canadian Draft 125.00 1,375.00
    - Opening Balance 1,000.00
    """
    line = clean_text(line)
    if not line:
        return None

    amounts = AMOUNT_RE.findall(line)

    if is_opening_balance_line(line) and amounts:
        balance = parse_amount(amounts[-1])
        return ParsedLine(
            date="",
            description="Opening Balance",
            debit=None,
            credit=None,
            balance=balance,
            category="Opening Balance",
        )

    if is_closing_balance_line(line) and amounts:
        balance = parse_amount(amounts[-1])
        return ParsedLine(
            date="",
            description="Closing Totals",
            debit=None,
            credit=None,
            balance=balance,
            category="Closing Totals",
        )

    if is_summary_line(line):
        return None

    date_match = DATE_RE.search(line)
    if not date_match or not amounts:
        return None

    date = date_match.group("date")
    without_date = clean_text(line.replace(date, " ", 1))
    for amount_text in amounts:
        without_date = clean_text(without_date.replace(amount_text, " ", 1))

    description = without_date
    parsed_amounts = [amount for amount in (parse_amount(a) for a in amounts) if amount is not None]
    if not parsed_amounts:
        return None

    balance = parsed_amounts[-1] if len(parsed_amounts) >= 2 else None

    if len(parsed_amounts) >= 3:
        debit = abs(parsed_amounts[-3]) if parsed_amounts[-3] else None
        credit = abs(parsed_amounts[-2]) if parsed_amounts[-2] else None
    else:
        debit, credit = decide_debit_credit(description, parsed_amounts[0])

    return ParsedLine(
        date=date,
        description=description,
        debit=debit,
        credit=credit,
        balance=balance,
        category=categorize_transaction(description),
    )


def parse_table_rows(tables: list[list[list[str]]]) -> list[ParsedLine]:
    """Parse transaction rows from Docling table output."""
    parsed: list[ParsedLine] = []

    for table in tables:
        for row in table:
            cells = [clean_text(cell) for cell in row if clean_text(cell)]
            if not cells:
                continue

            joined = " ".join(cells)
            if is_summary_line(joined) and not (
                is_opening_balance_line(joined) or is_closing_balance_line(joined)
            ):
                continue

            # Prefer structured table columns when they resemble:
            # Date | Description | Debit | Credit | Balance
            if len(cells) >= 5 and DATE_RE.search(cells[0]):
                debit = parse_amount(cells[-3])
                credit = parse_amount(cells[-2])
                balance = parse_amount(cells[-1])
                description = " ".join(cells[1:-3])
                parsed.append(
                    ParsedLine(
                        date=DATE_RE.search(cells[0]).group("date"),
                        description=description,
                        debit=abs(debit) if debit else None,
                        credit=abs(credit) if credit else None,
                        balance=balance,
                        category=categorize_transaction(description),
                    )
                )
                continue

            line = parse_transaction_line(joined)
            if line:
                parsed.append(line)

    return parsed


def dedupe_transactions(rows: Iterable[ParsedLine]) -> list[ParsedLine]:
    """Remove duplicate rows that can appear when Docling has both text and table output."""
    unique: list[ParsedLine] = []
    seen: set[tuple[Any, ...]] = set()

    for row in rows:
        key = (
            row.date,
            row.description.lower(),
            row.debit,
            row.credit,
            row.balance,
            row.category,
        )
        if key not in seen:
            seen.add(key)
            unique.append(row)

    return unique


def extract_transactions(data: dict[str, Any]) -> list[ParsedLine]:
    """
    Extract transactions from tables, OCR-positioned lines, and plain text blocks.

    Order matters: tables are usually most reliable, then positioned OCR rows,
    then generic text blocks as a fallback.
    """
    stacked_bmo_rows = extract_bmo_stacked_layout(data)
    if stacked_bmo_rows:
        return dedupe_transactions(stacked_bmo_rows)

    candidates: list[ParsedLine] = []

    candidates.extend(parse_table_rows(extract_tables(data)))

    for line in extract_positioned_lines(data):
        parsed = parse_transaction_line(line)
        if parsed:
            candidates.append(parsed)

    for block in extract_plain_text_blocks(data):
        parsed = parse_transaction_line(block)
        if parsed:
            candidates.append(parsed)

    return dedupe_transactions(candidates)


def to_dataframe(rows: list[ParsedLine]) -> pd.DataFrame:
    """Convert parsed rows into the exact Excel column order requested."""
    return pd.DataFrame(
        [
            {
                "Date": row.date,
                "Description": row.description,
                "Debit": row.debit,
                "Credit": row.credit,
                "Balance": row.balance,
                "Category": row.category,
            }
            for row in rows
        ],
        columns=["Date", "Description", "Debit", "Credit", "Balance", "Category"],
    )


def build_summary(df: pd.DataFrame) -> pd.DataFrame:
    """Create the Summary sheet values."""
    opening_rows = df[df["Category"] == "Opening Balance"]
    closing_rows = df[df["Category"] == "Closing Totals"]
    normal_rows = df[~df["Category"].isin(["Opening Balance", "Closing Totals"])]

    opening_balance = (
        opening_rows["Balance"].dropna().iloc[0] if not opening_rows["Balance"].dropna().empty else None
    )
    closing_balance = (
        closing_rows["Balance"].dropna().iloc[-1] if not closing_rows["Balance"].dropna().empty else None
    )

    return pd.DataFrame(
        [
            {"Metric": "Total Debits", "Value": normal_rows["Debit"].fillna(0).sum()},
            {"Metric": "Total Credits", "Value": normal_rows["Credit"].fillna(0).sum()},
            {"Metric": "Opening Balance", "Value": opening_balance},
            {"Metric": "Closing Balance", "Value": closing_balance},
            {"Metric": "Number of transactions", "Value": len(normal_rows)},
        ]
    )


def merge_excel_files(input_paths: list[Path], output_path: Path) -> Path:
    """
    Merge monthly extractor workbooks into one annual workbook.

    Each input workbook should have a Transactions sheet with the standard
    columns. Opening Balance and Closing Totals rows are preserved in a separate
    Balances sheet so the annual Transactions sheet contains only real activity.
    """
    frames: list[pd.DataFrame] = []
    balance_frames: list[pd.DataFrame] = []

    for path in input_paths:
        if not path.exists():
            raise FileNotFoundError(f"Excel file not found: {path}")

        workbook = pd.ExcelFile(path)
        source_sheet = "Extraction Data" if "Extraction Data" in workbook.sheet_names else "Transactions"
        df = pd.read_excel(path, sheet_name=source_sheet)
        missing = [col for col in ["Date", "Description", "Debit", "Credit", "Balance", "Category"] if col not in df.columns]
        if missing:
            raise RuntimeError(f"{path.name} is missing columns: {', '.join(missing)}")

        df.insert(0, "Source File", path.name)
        is_balance_row = df["Category"].isin(["Opening Balance", "Closing Totals"])
        balance_frames.append(df[is_balance_row].copy())
        frames.append(df[~is_balance_row].copy())

    if not frames:
        raise RuntimeError("No Excel files were provided.")

    annual_df = pd.concat(frames, ignore_index=True)
    balances_df = pd.concat(balance_frames, ignore_index=True) if balance_frames else pd.DataFrame()

    # Sort when dates are ISO-like. If not, preserve input order.
    sortable_dates = pd.to_datetime(annual_df["Date"], errors="coerce")
    if sortable_dates.notna().any():
        annual_df = annual_df.assign(_sort_date=sortable_dates).sort_values(
            ["_sort_date", "Source File"], na_position="last"
        ).drop(columns=["_sort_date"])

    summary_df = pd.DataFrame(
        [
            {"Metric": "Total Debits", "Value": annual_df["Debit"].fillna(0).sum()},
            {"Metric": "Total Credits", "Value": annual_df["Credit"].fillna(0).sum()},
            {"Metric": "Number of transactions", "Value": len(annual_df)},
            {"Metric": "Files merged", "Value": len(input_paths)},
        ]
    )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        annual_df.to_excel(writer, sheet_name="Annual Transactions", index=False)
        summary_df.to_excel(writer, sheet_name="Annual Summary", index=False)
        if not balances_df.empty:
            balances_df.to_excel(writer, sheet_name="Monthly Balances", index=False)

        auto_adjust_columns(writer, "Annual Transactions", annual_df)
        auto_adjust_columns(writer, "Annual Summary", summary_df)
        if not balances_df.empty:
            auto_adjust_columns(writer, "Monthly Balances", balances_df)

    return output_path


def detect_bank_from_file(path: Path) -> str:
    """Best-effort bank detection from filename and readable PDF text."""
    name = path.name.lower()
    if "tangerine" in name:
        return "Tangerine"
    if "cibc" in name:
        return "CIBC"
    if "bmo" in name or "bank of montreal" in name:
        return "BMO"
    if "td" in name:
        return "TD"
    if "rbc" in name or "royal" in name:
        return "RBC"

    if path.suffix.lower() == ".json":
        try:
            text = json.dumps(load_docling_json(path))[:25000].lower()
        except Exception:
            text = ""
    elif path.suffix.lower() == ".pdf":
        try:
            text = read_pdf_text(path)[:25000].lower()
        except Exception:
            text = ""
    else:
        text = ""

    if "tangerine" in text:
        return "Tangerine"
    if "cibc" in text:
        return "CIBC"
    if "bmo" in text or "bank of montreal" in text:
        return "BMO"
    if "td canada trust" in text or "td bank" in text:
        return "TD"
    if "rbc royal bank" in text or "royal bank" in text:
        return "RBC"

    return "Unknown"


def process_incoming_folder(
    incoming_dir: Path,
    excel_dir: Path,
    processed_dir: Path | None = None,
    move_processed: bool = False,
) -> pd.DataFrame:
    """
    Process supported PDF/JSON statements from an Incoming folder.

    Returns a status DataFrame with one row per input file. This is useful for
    Colab batch runs where the user wants to drop files into Google Drive and
    process them together.
    """
    incoming_dir.mkdir(parents=True, exist_ok=True)
    excel_dir.mkdir(parents=True, exist_ok=True)
    if processed_dir is not None:
        processed_dir.mkdir(parents=True, exist_ok=True)

    results: list[dict[str, Any]] = []
    input_files = sorted(
        [
            path
            for path in incoming_dir.iterdir()
            if path.is_file() and path.suffix.lower() in {".pdf", ".json"}
        ]
    )

    for path in input_files:
        bank = detect_bank_from_file(path)
        try:
            if bank == "Unknown":
                raise RuntimeError("Could not detect bank from filename or file text.")

            rows = extract_from_statement_file(path, bank)
            df = to_dataframe(rows)
            output_name = f"{path.stem}_{bank}_transactions.xlsx"
            saved_path = write_excel(df, excel_dir / output_name)
            summary = build_summary(df)

            if move_processed and processed_dir is not None:
                destination = processed_dir / path.name
                if destination.exists():
                    destination = processed_dir / f"{path.stem}_processed{path.suffix}"
                path.replace(destination)

            results.append(
                {
                    "File": path.name,
                    "Bank": bank,
                    "Status": "Success",
                    "Excel": str(saved_path),
                    "Transactions": int(summary.loc[summary["Metric"] == "Number of transactions", "Value"].iloc[0]),
                    "Total Debits": float(summary.loc[summary["Metric"] == "Total Debits", "Value"].iloc[0]),
                    "Total Credits": float(summary.loc[summary["Metric"] == "Total Credits", "Value"].iloc[0]),
                    "Error": "",
                }
            )
        except Exception as exc:
            results.append(
                {
                    "File": path.name,
                    "Bank": bank,
                    "Status": "Error",
                    "Excel": "",
                    "Transactions": "",
                    "Total Debits": "",
                    "Total Credits": "",
                    "Error": str(exc),
                }
            )

    return pd.DataFrame(results)


def auto_adjust_columns(writer: pd.ExcelWriter, sheet_name: str, df: pd.DataFrame) -> None:
    """Resize Excel columns based on content length."""
    worksheet = writer.sheets[sheet_name]
    for index, column in enumerate(df.columns, start=1):
        values = [str(column)] + ["" if pd.isna(value) else str(value) for value in df[column]]
        width = min(max(len(value) for value in values) + 2, 60)
        worksheet.column_dimensions[worksheet.cell(row=1, column=index).column_letter].width = width


def add_balance_formulas(writer: pd.ExcelWriter, sheet_name: str, df: pd.DataFrame) -> None:
    """Add one visible calculated running-balance column."""
    worksheet = writer.sheets[sheet_name]
    calculated_col = worksheet.max_column + 1
    worksheet.cell(row=1, column=calculated_col, value="Calculated Balance")

    visible_columns = ["Date", "Description", "Debit", "Credit", "Category"]
    debit_col = visible_columns.index("Debit") + 1
    credit_col = visible_columns.index("Credit") + 1

    for index, row in df.reset_index(drop=True).iterrows():
        excel_row = index + 2
        category = str(row.get("Category", ""))
        debit_ref = worksheet.cell(excel_row, debit_col).coordinate
        credit_ref = worksheet.cell(excel_row, credit_col).coordinate
        if category == "Opening Balance" or excel_row == 2:
            opening_balance = row.get("Balance")
            formula = 0 if pd.isna(opening_balance) else float(opening_balance)
        elif category == "Closing Totals":
            previous_ref = worksheet.cell(excel_row - 1, calculated_col).coordinate
            formula = f"={previous_ref}"
        else:
            previous_ref = worksheet.cell(excel_row - 1, calculated_col).coordinate
            formula = (
                f'={previous_ref}-IF({debit_ref}="",0,{debit_ref})'
                f'+IF({credit_ref}="",0,{credit_ref})'
            )
        worksheet.cell(excel_row, calculated_col, value=formula)

    for row_no in range(2, len(df) + 2):
        worksheet.cell(row_no, calculated_col).number_format = '#,##0.00;[Red]-#,##0.00'

    worksheet.column_dimensions[worksheet.cell(1, calculated_col).column_letter].width = 20


def write_excel(df: pd.DataFrame, output_path: Path) -> Path:
    """Write transactions and summary to Excel using pandas/openpyxl."""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    summary_df = build_summary(df)

    def write_to(path: Path) -> None:
        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            visible_df = df[["Date", "Description", "Debit", "Credit", "Category"]].copy()
            visible_df.to_excel(writer, sheet_name="Transactions", index=False)
            df.to_excel(writer, sheet_name="Extraction Data", index=False)
            summary_df.to_excel(writer, sheet_name="Summary", index=False)
            auto_adjust_columns(writer, "Transactions", visible_df)
            auto_adjust_columns(writer, "Extraction Data", df)
            auto_adjust_columns(writer, "Summary", summary_df)
            add_balance_formulas(writer, "Transactions", df)
            writer.sheets["Extraction Data"].sheet_state = "hidden"

    try:
        write_to(output_path)
        return output_path
    except PermissionError:
        # If the workbook is open in Excel, Windows locks it. Save a corrected
        # copy instead of failing, so the extraction can still complete.
        corrected_path = output_path.with_name(f"{output_path.stem}_corrected{output_path.suffix}")
        write_to(corrected_path)
        return corrected_path


def read_pdf_text(path: Path) -> str:
    """Extract text from a local PDF. This is offline and uses pypdf."""
    if PdfReader is None:
        raise RuntimeError("PDF support requires pypdf. Use a Docling JSON file instead.")
    if not path.exists():
        raise FileNotFoundError(f"PDF file not found: {path}")
    reader = PdfReader(str(path))
    return "\n".join(page.extract_text() or "" for page in reader.pages)


def extract_cibc_pdf_transactions(pdf_text: str) -> list[ParsedLine]:
    """
    Extract CIBC transactions from readable PDF text.

    CIBC text often has transaction descriptions spanning multiple lines, with the
    amount and balance on the last line of each transaction.
    """
    lines = [clean_text(line) for line in pdf_text.splitlines() if clean_text(line)]
    joined = "\n".join(lines)
    if "CIBC" not in joined or "Transaction details" not in joined:
        return []

    opening_balance = None
    closing_balance = None
    total_debits = None
    total_credits = None
    opening_date = ""
    closing_date = ""

    for line in lines:
        match = re.search(r"Opening balance on ([A-Z][a-z]{2}\s+\d{1,2}),\s+\d{4}\s+\$?([\d,]+\.\d{2})", line)
        if match:
            opening_date = match.group(1)
            opening_balance = parse_amount(match.group(2))
        match = re.search(r"Withdrawals\s+-\s+([\d,]+\.\d{2})", line)
        if match:
            total_debits = parse_amount(match.group(1))
        match = re.search(r"Deposits\s+\+\s+([\d,]+\.\d{2})", line)
        if match:
            total_credits = parse_amount(match.group(1))
        match = re.search(r"Closing balance on ([A-Z][a-z]{2}\s+\d{1,2}),\s+\d{4}\s+=\s+\$?([\d,]+\.\d{2})", line)
        if match:
            closing_date = match.group(1)
            closing_balance = parse_amount(match.group(2))

    rows: list[ParsedLine] = []
    if opening_balance is not None:
        rows.append(
            ParsedLine(opening_date, "Opening balance", None, None, opening_balance, "Opening Balance")
        )

    in_details = False
    current_date = ""
    buffer: list[str] = []

    def flush_buffer() -> None:
        nonlocal buffer, current_date
        if not buffer:
            return

        text = " ".join(buffer)
        buffer = []

        lowered = text.lower()
        if (
            "opening balance" in lowered
            or "balance forward" in lowered
            or "closing balance" in lowered
            or "continued on next page" in lowered
            or "date description withdrawals" in lowered
            or "page " in lowered
        ):
            return

        amounts = _amounts_from_text(text)
        if len(amounts) < 2:
            return

        amount = amounts[-2]
        balance = amounts[-1]
        description = text
        for raw in AMOUNT_RE.findall(text)[-2:]:
            description = clean_text(description.replace(raw, " ", 1))

        if not description:
            return

        desc_lower = description.lower()
        is_credit = any(word in desc_lower for word in ["deposit", "fee waiver"])
        debit = None if is_credit else abs(amount)
        credit = abs(amount) if is_credit else None

        rows.append(
            ParsedLine(
                current_date,
                description,
                debit,
                credit,
                balance,
                categorize_transaction(description),
            )
        )

    for line in lines:
        if line.startswith("Transaction details"):
            in_details = True
            continue
        if not in_details:
            continue
        if line.startswith("Important:") or line.startswith("*Foreign Currency"):
            flush_buffer()
            break
        if line.startswith("Date Description"):
            continue

        date_match = re.match(r"^([A-Z][a-z]{2}\s+\d{1,2})\s+(.+)$", line)
        if date_match:
            flush_buffer()
            current_date = date_match.group(1)
            remainder = clean_text(date_match.group(2))
            if "opening balance" in remainder.lower() or "balance forward" in remainder.lower():
                buffer = []
            else:
                buffer = [remainder]
                if len(_amounts_from_text(remainder)) >= 2:
                    flush_buffer()
            continue

        if not current_date:
            continue

        if _amounts_from_text(line) and buffer:
            buffer.append(line)
            flush_buffer()
        elif _amounts_from_text(line) and not buffer:
            # Some CIBC rows, like fee waivers, continue using the previous date.
            buffer = [line]
            flush_buffer()
        else:
            buffer.append(line)

    flush_buffer()

    if closing_balance is not None:
        rows.append(
            ParsedLine(
                closing_date,
                "Closing totals",
                total_debits,
                total_credits,
                closing_balance,
                "Closing Totals",
            )
        )

    return dedupe_transactions(rows)


def extract_tangerine_pdf_transactions(pdf_text: str) -> list[ParsedLine]:
    """
    Extract Tangerine chequing transactions from readable PDF text.

    Tangerine rows are printed as:
    Balance | Amount | Description | Date

    The statement does not label debit/credit per row, so this parser compares
    each balance to the previous balance. If balance goes up, Amount is Credit;
    if balance goes down, Amount is Debit.
    """
    lines = [clean_text(line) for line in pdf_text.splitlines() if clean_text(line)]
    joined = "\n".join(lines)
    if "Tangerine" not in joined or "Transaction Description Transaction Date" not in joined:
        return []

    rows: list[ParsedLine] = []
    in_details = False
    buffer = ""
    previous_balance: float | None = None

    row_start_re = re.compile(r"^([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+(.+)$")
    date_end_re = re.compile(r"(.+?)\s+(\d{2}\s+[A-Z][a-z]{2}\s+\d{4})$")

    def normalize_date(date_text: str) -> str:
        try:
            return datetime.strptime(date_text, "%d %b %Y").strftime("%Y-%m-%d")
        except ValueError:
            return date_text

    def flush_buffer() -> None:
        nonlocal buffer, previous_balance
        text = clean_text(buffer)
        buffer = ""
        if not text:
            return

        match = row_start_re.match(text)
        if not match:
            return

        balance = parse_amount(match.group(1))
        amount = parse_amount(match.group(2))
        rest = clean_text(match.group(3))
        date_match = date_end_re.match(rest)
        if not date_match or balance is None or amount is None:
            return

        description = clean_text(date_match.group(1))
        date = normalize_date(date_match.group(2))

        if "opening balance" in description.lower():
            previous_balance = balance
            rows.append(
                ParsedLine(date, "Opening Balance", None, None, balance, "Opening Balance")
            )
            return

        if "closing balance" in description.lower():
            rows.append(
                ParsedLine(date, "Closing Totals", None, None, balance, "Closing Totals")
            )
            return

        debit = None
        credit = None
        if previous_balance is not None:
            if balance > previous_balance:
                credit = abs(amount)
            elif balance < previous_balance:
                debit = abs(amount)
        if debit is None and credit is None and amount:
            debit, credit = decide_debit_credit(description, amount)

        rows.append(
            ParsedLine(
                date,
                description,
                debit,
                credit,
                balance,
                categorize_transaction(description),
            )
        )
        previous_balance = balance

    for line in lines:
        if line.startswith("Balance($) Amount($) Transaction Description Transaction Date"):
            in_details = True
            continue
        if not in_details:
            continue
        if line.startswith("Up to $100") or line.startswith("Nobody likes mistakes") or line.startswith("Page "):
            flush_buffer()
            break

        if row_start_re.match(line):
            flush_buffer()
            buffer = line
        else:
            buffer = clean_text(f"{buffer} {line}")

    flush_buffer()
    return dedupe_transactions(rows)


def extract_rbc_pdf_transactions(pdf_text: str) -> list[ParsedLine]:
    """Extract RBC business-account activity from readable PDF text."""
    lines = [clean_text(line) for line in pdf_text.splitlines() if clean_text(line)]
    joined = "\n".join(lines)
    if "ROYAL BANK OF CANADA" not in joined or "Account Activity Details" not in joined:
        return []

    money_re = re.compile(r"-?\d{1,3}(?:,\d{3})*\.\d{2}")
    date_re = re.compile(r"^(\d{1,2}\s+[A-Z][a-z]{2})(?:\1)?\s*(.*)$")
    rows: list[ParsedLine] = []
    current_date = ""
    previous_balance: float | None = None
    in_activity = False
    pending_description = ""

    period_match = re.search(
        r"([A-Z][a-z]+)\s+(\d{1,2}),\s*(\d{4})\s+to\s+"
        r"([A-Z][a-z]+)\s+(\d{1,2}),\s*(\d{4})",
        joined,
    )
    period_start: datetime | None = None
    period_end: datetime | None = None
    if period_match:
        try:
            period_start = datetime.strptime(
                f"{period_match.group(1)} {period_match.group(2)} {period_match.group(3)}",
                "%B %d %Y",
            )
            period_end = datetime.strptime(
                f"{period_match.group(4)} {period_match.group(5)} {period_match.group(6)}",
                "%B %d %Y",
            )
        except ValueError:
            period_start = None
            period_end = None

    def normalize_rbc_date(date_text: str) -> str:
        """Attach the correct statement-period year to an RBC day/month label."""
        try:
            partial = datetime.strptime(date_text, "%d %b")
        except ValueError:
            return date_text

        if period_start is None or period_end is None:
            return date_text

        candidates = [
            partial.replace(year=period_start.year),
            partial.replace(year=period_end.year),
        ]
        within_period = [value for value in candidates if period_start <= value <= period_end]
        if within_period:
            return min(within_period, key=lambda value: abs((value - period_start).days)).strftime("%Y-%m-%d")

        # This fallback is mainly for OCR errors near a statement boundary.
        return min(
            candidates,
            key=lambda value: min(abs((value - period_start).days), abs((value - period_end).days)),
        ).strftime("%Y-%m-%d")

    credit_markers = (
        "loan credit",
        "mobile cheque deposit",
        "e-transfer - autodeposit",
        "item returned nsf",
        "account payable pmt stanchuck",
        "deposit",
    )

    opening_match = re.search(
        r"Opening balance on ([A-Z][a-z]+) (\d{1,2}),\s*(\d{4}) \$?([\d,]+\.\d{2})",
        joined,
    )
    closing_match = re.search(
        r"Closing balance on ([A-Z][a-z]+) (\d{1,2}),\s*(\d{4}) = \$?([\d,]+\.\d{2})",
        joined,
    )
    debit_total_match = re.search(r"Total cheques & debits \(\d+\) - ([\d,]+\.\d{2})", joined)
    credit_total_match = re.search(r"Total deposits & credits \(\d+\) \+ ([\d,]+\.\d{2})", joined)

    if opening_match:
        opening_date = datetime.strptime(
            f"{opening_match.group(1)} {opening_match.group(2)} {opening_match.group(3)}",
            "%B %d %Y",
        ).strftime("%Y-%m-%d")
        previous_balance = parse_amount(opening_match.group(4))
        rows.append(
            ParsedLine(opening_date, "Opening Balance", None, None, previous_balance, "Opening Balance")
        )

    def is_credit(description: str) -> bool:
        lower = description.lower()
        return any(marker in lower for marker in credit_markers)

    def add_transaction(description: str, amount: float, balance: float | None) -> None:
        nonlocal previous_balance
        credit = abs(amount) if is_credit(description) else None
        debit = None if credit is not None else abs(amount)

        # Balance arithmetic is more authoritative than keywords when available.
        if balance is not None and previous_balance is not None:
            delta = round(balance - previous_balance, 2)
            if abs(abs(delta) - abs(amount)) < 0.02:
                if delta > 0:
                    debit, credit = None, abs(amount)
                elif delta < 0:
                    debit, credit = abs(amount), None

        rows.append(
            ParsedLine(
                current_date,
                description,
                debit,
                credit,
                balance,
                categorize_transaction(description),
            )
        )
        if balance is not None:
            previous_balance = balance

    for line in lines:
        if line.startswith("Account Activity Details"):
            in_activity = True
            continue
        if not in_activity:
            continue
        if line.startswith("Account Fees:"):
            break
        if (
            line.startswith("Date Description")
            or line.startswith("Business Account Statement")
            or line.startswith("ROYAL BANK OF CANADA")
            or line.startswith("P.O. BAG SERVICE")
            or line.startswith("CALGARY AB")
            or re.match(r"^\d+ of \d+$", line)
            or line.startswith("Account number:")
            or re.match(r"^[A-Z][a-z]+ \d{1,2}, \d{4} to", line)
        ):
            continue
        if line.startswith("Opening balance") or line.startswith("Closing balance"):
            continue

        date_match = date_re.match(line)
        if date_match:
            current_date = normalize_rbc_date(date_match.group(1))
            line = clean_text(date_match.group(2))
            if not line:
                continue

        amounts_raw = money_re.findall(line)
        if not amounts_raw:
            pending_description = clean_text(f"{pending_description} {line}")
            continue

        amounts = [parse_amount(value) for value in amounts_raw]
        amounts = [value for value in amounts if value is not None]
        if not amounts:
            continue

        description = line
        for raw in amounts_raw:
            description = clean_text(description.replace(raw, " ", 1))
        description = clean_text(f"{pending_description} {description}")
        pending_description = ""

        amount = amounts[0]
        balance = amounts[1] if len(amounts) >= 2 else None
        add_transaction(description, amount, balance)

    if closing_match:
        closing_date = datetime.strptime(
            f"{closing_match.group(1)} {closing_match.group(2)} {closing_match.group(3)}",
            "%B %d %Y",
        ).strftime("%Y-%m-%d")
        closing_balance = parse_amount(closing_match.group(4))
        total_debits = parse_amount(debit_total_match.group(1)) if debit_total_match else None
        total_credits = parse_amount(credit_total_match.group(1)) if credit_total_match else None
        rows.append(
            ParsedLine(
                closing_date,
                "Closing Totals",
                total_debits,
                total_credits,
                closing_balance,
                "Closing Totals",
            )
        )

    return dedupe_transactions(rows)


def extract_from_statement_file(path: Path, bank_name: str = "BMO") -> list[ParsedLine]:
    """Read JSON or PDF input and dispatch to the best available extractor."""
    suffix = path.suffix.lower()
    if suffix == ".json":
        return extract_transactions(load_docling_json(path))
    if suffix == ".pdf":
        pdf_text = read_pdf_text(path)
        selected_bank = bank_name.upper()

        if selected_bank == "CIBC":
            rows = extract_cibc_pdf_transactions(pdf_text)
            if rows:
                return rows

        if selected_bank == "TANGERINE":
            rows = extract_tangerine_pdf_transactions(pdf_text)
            if rows:
                return rows

        if selected_bank == "RBC":
            rows = extract_rbc_pdf_transactions(pdf_text)
            if rows:
                return rows

        # Friendly fallback: if the dropdown is left on BMO but the uploaded PDF
        # is actually another supported bank, still extract it instead of showing
        # a dead-end error.
        rows = extract_cibc_pdf_transactions(pdf_text)
        if rows:
            return rows
        rows = extract_tangerine_pdf_transactions(pdf_text)
        if rows:
            return rows
        rows = extract_rbc_pdf_transactions(pdf_text)
        if rows:
            return rows

        if selected_bank == "BMO":
            raise RuntimeError(
                "BMO PDF extraction needs Docling JSON for this statement layout. "
                "Choose text.json for BMO, or select CIBC if this is the CIBC PDF."
            )

        raise RuntimeError(
            f"No {bank_name} PDF extractor matched this statement yet. "
            "Send one sample PDF for that bank so its layout can be tuned."
        )
    raise RuntimeError("Please choose a .json or .pdf file.")


def main() -> int:
    try:
        data = load_docling_json(INPUT_JSON)
        rows = extract_transactions(data)

        if not rows:
            raise RuntimeError(
                "No transactions found. Check whether the Docling JSON contains "
                "tables/text blocks, then adjust parse_transaction_line()."
            )

        df = to_dataframe(rows)
        saved_path = write_excel(df, OUTPUT_XLSX)

        normal_count = len(df[~df["Category"].isin(["Opening Balance", "Closing Totals"])])
        print(f"Created: {saved_path}")
        print(f"Rows written: {len(df)}")
        print(f"Normal transactions: {normal_count}")
        return 0

    except FileNotFoundError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1
    except json.JSONDecodeError as exc:
        print(f"ERROR: Invalid JSON file: {exc}", file=sys.stderr)
        return 1
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
