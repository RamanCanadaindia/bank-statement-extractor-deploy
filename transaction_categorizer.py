from __future__ import annotations

import io
import re
from dataclasses import dataclass

import pandas as pd


@dataclass(frozen=True)
class CategoryRule:
    keywords: tuple[str, ...]
    category: str
    applies_to: str = "Any"
    confidence: str = "High"


DEFAULT_RULES = [
    CategoryRule(("opening balance",), "Opening Balance"),
    CategoryRule(("closing balance", "closing totals"), "Closing Totals"),
    CategoryRule(("purchase interest", "cash advance interest", "interest charge"), "Interest Expense"),
    CategoryRule(("account fee", "monthly fee", "service charge", "draft fee", "nsf fee", "overlimit fee"), "Bank Charges"),
    CategoryRule(("icbc", "insurance", "intact insurance", "wawanesa", "aviva"), "Insurance"),
    CategoryRule(("shell", "petro-canada", "petro canada", "esso", "chevron", "husky", "fuel"), "Fuel", "Debit"),
    CategoryRule(("uber", "lyft", "taxi", "translink"), "Travel/Transportation", "Debit"),
    CategoryRule(("air canada", "westjet", "hotel", "expedia", "booking.com"), "Travel", "Debit"),
    CategoryRule(("tim hortons", "starbucks", "mcdonald", "restaurant", "doordash", "skipthe"), "Meals and Entertainment", "Debit"),
    CategoryRule(("telus", "rogers", "bell mobility", "freedom mobile", "fido"), "Telephone and Internet", "Debit"),
    CategoryRule(("bc hydro", "fortisbc", "fortis bc", "hydro bill"), "Utilities", "Debit"),
    CategoryRule(("staples", "office depot"), "Office Supplies", "Debit"),
    CategoryRule(("canada post", "fedex", "ups canada", "purolator"), "Postage and Courier", "Debit"),
    CategoryRule(("quickbooks", "intuit", "microsoft", "adobe", "dropbox", "google workspace"), "Software and Subscriptions", "Debit"),
    CategoryRule(("accounting fee", "legal fee", "law office", "consulting fee"), "Professional Fees", "Debit"),
    CategoryRule(("facebook ads", "meta ads", "google ads", "advertising"), "Advertising and Marketing", "Debit"),
    CategoryRule(("rent/lease", "lease payment", "commercial rent", "comm rent"), "Rent and Lease", "Debit"),
    CategoryRule(("repair", "maintenance"), "Repairs and Maintenance", "Debit"),
    CategoryRule(("cra payment", "canada revenue", "receiver general", "property tax"), "Taxes and Licences", "Debit"),
    CategoryRule(("payroll", "wage", "salary", "pay emp-vendor"), "Payroll and Wages", "Debit"),
    CategoryRule(("credit card payment", "payment - thank you", "paiement - merci"), "Credit Card Payment"),
    CategoryRule(("loan payment", "mortgage payment"), "Loan Payment", "Debit"),
    CategoryRule(("loan credit", "loan advance", "line of credit advance"), "Loan Proceeds", "Credit"),
    CategoryRule(("stripe", "square", "shopify payments"), "Sales Revenue", "Credit"),
    CategoryRule(("stripe fee", "square fee", "merchant fee"), "Merchant Processing Fees", "Debit"),
    CategoryRule(("tax refund", "cra refund"), "Tax Refund", "Credit"),
    CategoryRule(("interest deposit", "interest earned"), "Interest Income", "Credit"),
    CategoryRule(("customer payment", "invoice payment", "sales deposit"), "Sales Revenue", "Credit"),
    CategoryRule(("deposit", "mobile cheque deposit", "direct deposit"), "Income/Deposit", "Credit"),
    CategoryRule(("e-transfer sent", "etransfer sent"), "Outgoing Transfer", "Debit"),
    CategoryRule(("e-transfer received", "e-transfer autodeposit", "etransfer received"), "Incoming Transfer", "Credit"),
    CategoryRule(("online banking transfer", "account transfer", "funds transfer"), "Transfer", "Any", "Medium"),
    CategoryRule(("canadian draft", "bank draft"), "Bank Transfer/Payment"),
    CategoryRule(("pre-auth debit", "pre-authorized payment", "pad payment"), "Pre-Authorized Payment", "Debit"),
    CategoryRule(("bill payment", "internet bill pmt"), "Bill Payment", "Debit"),
    CategoryRule(("cheque", "check no"), "Cheque", "Debit"),
    CategoryRule(("costco", "walmart", "amazon"), "General Purchases", "Debit", "Low"),
]

SPECIAL_CATEGORIES = {"Opening Balance", "Closing Totals"}


def _clean_column(value: object) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(value).lower())


def _find_column(frame: pd.DataFrame, names: tuple[str, ...]) -> str | None:
    normalized = {_clean_column(column): str(column) for column in frame.columns}
    for name in names:
        if _clean_column(name) in normalized:
            return normalized[_clean_column(name)]
    return None


def load_transactions(payload: bytes, filename: str) -> pd.DataFrame:
    """Load an extractor workbook, ordinary Excel file, or CSV transaction list."""
    if filename.lower().endswith(".csv"):
        frame = pd.read_csv(io.BytesIO(payload))
    else:
        workbook = pd.ExcelFile(io.BytesIO(payload))
        preferred = next(
            (
                sheet
                for sheet in workbook.sheet_names
                if sheet.lower() in {"transactions", "categorized transactions", "annual transactions"}
            ),
            workbook.sheet_names[0],
        )
        frame = pd.read_excel(workbook, sheet_name=preferred)

    description_column = _find_column(
        frame,
        ("Description", "Transaction Description", "Details", "Memo", "Payee", "Merchant"),
    )
    if description_column is None:
        raise ValueError("No Description, Details, Memo, Payee, or Merchant column was found.")

    result = frame.copy()
    if description_column != "Description":
        result = result.rename(columns={description_column: "Description"})
    result["Description"] = result["Description"].fillna("").astype(str).str.strip()

    amount_column = _find_column(result, ("Amount", "Transaction Amount", "Net Amount"))
    debit_column = _find_column(result, ("Debit", "Withdrawal", "Withdrawals", "Charges"))
    credit_column = _find_column(result, ("Credit", "Deposit", "Deposits", "Payments"))
    if amount_column:
        result["Amount"] = pd.to_numeric(result[amount_column], errors="coerce")
    elif debit_column or credit_column:
        debits = (
            pd.to_numeric(result[debit_column], errors="coerce").fillna(0)
            if debit_column
            else pd.Series(0.0, index=result.index)
        )
        credits = (
            pd.to_numeric(result[credit_column], errors="coerce").fillna(0)
            if credit_column
            else pd.Series(0.0, index=result.index)
        )
        result["Amount"] = credits.abs() - debits.abs()
    else:
        raise ValueError("No Amount column or Debit/Credit columns were found.")

    return result


def parse_custom_rules(frame: pd.DataFrame | None) -> list[CategoryRule]:
    if frame is None or frame.empty:
        return []
    keyword_column = _find_column(frame, ("Keyword", "Keywords", "Merchant", "Description Contains"))
    category_column = _find_column(frame, ("Category",))
    applies_column = _find_column(frame, ("Applies To", "Type", "Direction"))
    if keyword_column is None or category_column is None:
        raise ValueError("Custom rules require Keyword and Category columns.")

    rules: list[CategoryRule] = []
    for row in frame.to_dict("records"):
        keyword = str(row.get(keyword_column, "")).strip().lower()
        category = str(row.get(category_column, "")).strip()
        if not keyword or not category or keyword == "nan" or category == "nan":
            continue
        applies_to = str(row.get(applies_column, "Any")).strip().title() if applies_column else "Any"
        if applies_to not in {"Any", "Debit", "Credit"}:
            applies_to = "Any"
        rules.append(CategoryRule((keyword,), category, applies_to, "Custom"))
    return rules


def categorize_description(
    description: str,
    amount: float | None,
    custom_rules: list[CategoryRule] | None = None,
) -> tuple[str, str, str]:
    desc = re.sub(r"\s+", " ", str(description).lower()).strip()
    direction = "Credit" if pd.notna(amount) and float(amount) > 0 else "Debit"
    for rule in [*(custom_rules or []), *DEFAULT_RULES]:
        if rule.applies_to not in {"Any", direction}:
            continue
        matched = next((keyword for keyword in rule.keywords if keyword in desc), None)
        if matched:
            return rule.category, rule.confidence, matched

    if any(word in desc for word in ("fee", "charge")):
        return "Bank Charges", "Medium", "fee/charge fallback"
    return "Uncategorized", "Review", ""


def categorize_transactions(
    frame: pd.DataFrame,
    custom_rules: list[CategoryRule] | None = None,
    overwrite_existing: bool = False,
) -> pd.DataFrame:
    result = frame.copy()
    if "Category" not in result:
        result["Category"] = ""
    result["Category"] = result["Category"].fillna("").astype(str)
    result["Category Confidence"] = ""
    result["Category Rule"] = ""

    for index, row in result.iterrows():
        existing = str(row.get("Category", "")).strip()
        if existing in SPECIAL_CATEGORIES:
            result.at[index, "Category Confidence"] = "Protected"
            result.at[index, "Category Rule"] = "statement balance row"
            continue
        if existing and existing != "Uncategorized" and not overwrite_existing:
            result.at[index, "Category Confidence"] = "Existing"
            result.at[index, "Category Rule"] = "existing category retained"
            continue
        category, confidence, rule = categorize_description(
            row.get("Description", ""),
            row.get("Amount"),
            custom_rules,
        )
        result.at[index, "Category"] = category
        result.at[index, "Category Confidence"] = confidence
        result.at[index, "Category Rule"] = rule
    return result


def category_summary(frame: pd.DataFrame) -> pd.DataFrame:
    normal = frame[~frame["Category"].isin(SPECIAL_CATEGORIES)].copy()
    return (
        normal.groupby("Category", dropna=False)
        .agg(Transactions=("Description", "count"), Net_Amount=("Amount", "sum"))
        .reset_index()
        .rename(columns={"Net_Amount": "Net Amount"})
        .sort_values(["Category"])
        .reset_index(drop=True)
    )


def export_categorized_workbook(frame: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    summary = category_summary(frame)
    rules = pd.DataFrame(
        [
            {
                "Keywords": ", ".join(rule.keywords),
                "Category": rule.category,
                "Applies To": rule.applies_to,
                "Confidence": rule.confidence,
            }
            for rule in DEFAULT_RULES
        ]
    )
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        frame.to_excel(writer, sheet_name="Categorized Transactions", index=False)
        summary.to_excel(writer, sheet_name="Category Summary", index=False)
        rules.to_excel(writer, sheet_name="Built-in Rules", index=False)
        for sheet_name in writer.book.sheetnames:
            sheet = writer.book[sheet_name]
            sheet.freeze_panes = "A2"
            sheet.auto_filter.ref = sheet.dimensions
            for column_cells in sheet.columns:
                width = max(len(str(cell.value or "")) for cell in column_cells)
                sheet.column_dimensions[column_cells[0].column_letter].width = min(max(width + 2, 11), 48)
    return output.getvalue()


def custom_rule_template() -> bytes:
    output = io.BytesIO()
    pd.DataFrame(
        [
            {"Keyword": "merchant name", "Category": "Office Supplies", "Applies To": "Debit"},
            {"Keyword": "customer name", "Category": "Sales Revenue", "Applies To": "Credit"},
        ]
    ).to_csv(output, index=False)
    return output.getvalue()
