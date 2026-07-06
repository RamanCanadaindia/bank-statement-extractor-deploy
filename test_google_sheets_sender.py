import json
import unittest
from unittest.mock import patch

import pandas as pd

import google_sheets_sender as sender


class FakeResponse:
    def __init__(self, value):
        self.value = value

    def __enter__(self):
        return self

    def __exit__(self, *args):
        return False

    def read(self):
        return json.dumps(self.value).encode("utf-8")


class GoogleSheetsSenderTests(unittest.TestCase):
    def test_accepts_sheet_url_and_serializes_transactions(self):
        frame = pd.DataFrame(
            [{"Date": pd.Timestamp("2026-01-02"), "Description": "Deposit", "Amount": 100.0}]
        )
        with patch(
            "google_sheets_sender.urllib.request.urlopen",
            return_value=FakeResponse({"ok": True, "rows_added": 1}),
        ) as urlopen:
            result = sender.send_transactions(
                endpoint_url="https://script.google.com/macros/s/ABC_123/exec",
                shared_secret="private",
                spreadsheet="https://docs.google.com/spreadsheets/d/1nv4k-mRjNQ_2U5uuMS8z26ZQWHMyuILR1cndIAgs0mk/edit",
                sheet_name="Bank Transactions",
                frame=frame,
                source_file="January.pdf",
                bank="RBC",
            )

        self.assertEqual(result["rows_added"], 1)
        request = urlopen.call_args.args[0]
        payload = json.loads(request.data.decode("utf-8"))
        self.assertEqual(payload["rows"][0]["Date"], "2026-01-02T00:00:00")
        self.assertEqual(payload["rows"][0]["Source File"], "January.pdf")
        self.assertEqual(payload["rows"][0]["Bank"], "RBC")
        self.assertEqual(len(payload["batch_id"]), 64)

    def test_rejects_non_google_endpoint(self):
        with self.assertRaisesRegex(ValueError, "Google Apps Script"):
            sender.normalize_script_url("https://example.com/collect")

    def test_batch_id_is_stable(self):
        rows = [{"Description": "Fee", "Amount": -5.0}]
        first = sender.build_batch_id("one.pdf", "BMO", rows)
        second = sender.build_batch_id("one.pdf", "BMO", rows)
        self.assertEqual(first, second)


if __name__ == "__main__":
    unittest.main()
