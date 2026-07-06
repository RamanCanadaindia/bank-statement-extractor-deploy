"""Send reviewed transaction rows to a Google Apps Script Web App."""

from __future__ import annotations

import hashlib
import json
import math
import re
import urllib.error
import urllib.parse
import urllib.request
from datetime import date, datetime
from typing import Any

import pandas as pd


SCRIPT_URL_RE = re.compile(r"^https://script\.google\.com/macros/s/[A-Za-z0-9_-]+/exec$")
SPREADSHEET_ID_RE = re.compile(r"^[A-Za-z0-9_-]{20,}$")


def normalize_script_url(value: str) -> str:
    """Accept only a deployed Google Apps Script HTTPS endpoint."""
    url = value.strip()
    parsed = urllib.parse.urlparse(url)
    normalized = urllib.parse.urlunparse(
        (parsed.scheme, parsed.netloc, parsed.path.rstrip("/"), "", "", "")
    )
    if not SCRIPT_URL_RE.fullmatch(normalized):
        raise ValueError(
            "Enter a deployed Google Apps Script URL ending in /exec."
        )
    return normalized


def extract_spreadsheet_id(value: str) -> str:
    """Accept either a Google Sheets URL or its spreadsheet ID."""
    candidate = value.strip()
    match = re.search(r"/spreadsheets/d/([A-Za-z0-9_-]+)", candidate)
    spreadsheet_id = match.group(1) if match else candidate
    if not SPREADSHEET_ID_RE.fullmatch(spreadsheet_id):
        raise ValueError("Enter a valid Google Sheets URL or spreadsheet ID.")
    return spreadsheet_id


def _json_value(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, (datetime, date, pd.Timestamp)):
        return value.isoformat()
    if hasattr(value, "item"):
        value = value.item()
    if isinstance(value, float) and math.isnan(value):
        return None
    return value


def prepare_rows(frame: pd.DataFrame, source_file: str, bank: str) -> list[dict[str, Any]]:
    """Convert a transaction frame into JSON-safe records with audit fields."""
    prepared = frame.copy()
    if "Source File" not in prepared.columns:
        prepared["Source File"] = source_file
    if "Bank" not in prepared.columns:
        prepared["Bank"] = bank
    return [
        {str(column): _json_value(value) for column, value in row.items()}
        for row in prepared.to_dict(orient="records")
    ]


def build_batch_id(source_file: str, bank: str, rows: list[dict[str, Any]]) -> str:
    """Create a stable identifier so the receiver can reject duplicate uploads."""
    canonical = json.dumps(
        {"source_file": source_file, "bank": bank, "rows": rows},
        sort_keys=True,
        separators=(",", ":"),
        ensure_ascii=True,
    )
    return hashlib.sha256(canonical.encode("utf-8")).hexdigest()


def send_transactions(
    *,
    endpoint_url: str,
    shared_secret: str,
    spreadsheet: str,
    sheet_name: str,
    frame: pd.DataFrame,
    source_file: str,
    bank: str,
    timeout: int = 45,
) -> dict[str, Any]:
    """POST one statement's transactions to the Apps Script receiver."""
    endpoint = normalize_script_url(endpoint_url)
    spreadsheet_id = extract_spreadsheet_id(spreadsheet)
    target_sheet = sheet_name.strip()
    if not target_sheet:
        raise ValueError("Enter the destination tab name.")
    if not shared_secret:
        raise ValueError("Enter the connector secret.")

    rows = prepare_rows(frame, source_file, bank)
    if not rows:
        raise ValueError("There are no transactions to send.")

    payload = {
        "action": "append_transactions",
        "secret": shared_secret,
        "spreadsheet_id": spreadsheet_id,
        "sheet_name": target_sheet,
        "source_file": source_file,
        "bank": bank,
        "batch_id": build_batch_id(source_file, bank, rows),
        "rows": rows,
    }
    request = urllib.request.Request(
        endpoint,
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json"},
        method="POST",
    )

    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            result = json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"Google Sheets connector returned HTTP {exc.code}: {detail}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"Could not reach the Google Sheets connector: {exc.reason}") from exc
    except json.JSONDecodeError as exc:
        raise RuntimeError("The Google Sheets connector returned an invalid response.") from exc

    if not result.get("ok"):
        raise RuntimeError(result.get("error") or "Google Sheets rejected the transaction upload.")
    return result
