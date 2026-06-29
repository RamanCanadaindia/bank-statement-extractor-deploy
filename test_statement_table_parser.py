import unittest

import bmo_docling_to_excel as extractor


class StatementTableParserTests(unittest.TestCase):
    def test_blank_cells_preserve_debit_and_credit_columns(self):
        rows = extractor.parse_table_rows(
            [
                [
                    ["Jun 30, 2025", "MONTHLY ACCOUNT FEE", "10.95", "", "7,801.21"],
                    ["Jun 27, 2025", "TD ATM DEP 00405", "", "1,200.00", "8,654.44"],
                ]
            ]
        )

        self.assertEqual(rows[0].debit, 10.95)
        self.assertIsNone(rows[0].credit)
        self.assertIsNone(rows[1].debit)
        self.assertEqual(rows[1].credit, 1200.00)

    def test_positioned_ocr_preserves_scanned_table_columns(self):
        words = [
            ("Date", 20, 100),
            ("Description", 150, 100),
            ("Debit", 500, 100),
            ("Credit", 620, 100),
            ("Balance", 750, 100),
            ("Jun", 20, 150),
            ("30,", 60, 150),
            ("2025", 95, 150),
            ("MONTHLY", 170, 150),
            ("ACCOUNT", 260, 150),
            ("FEE", 355, 150),
            ("10.95", 505, 150),
            ("7,801.21", 745, 150),
            ("Jun", 20, 190),
            ("27,", 60, 190),
            ("2025", 95, 190),
            ("TD", 170, 190),
            ("ATM", 205, 190),
            ("DEP", 250, 190),
            ("00405", 300, 190),
            ("1,200.00", 615, 190),
            ("8,654.44", 745, 190),
        ]
        data = {
            "text": [word for word, _, _ in words],
            "left": [x for _, x, _ in words],
            "top": [y for _, _, y in words],
            "width": [55] * len(words),
            "height": [20] * len(words),
            "conf": [95] * len(words),
        }

        rows = extractor.parse_scanned_debit_credit_ocr_data(data)

        self.assertEqual(rows[0].date, "2025-06-30")
        self.assertEqual(rows[0].debit, 10.95)
        self.assertIsNone(rows[0].credit)
        self.assertEqual(rows[1].date, "2025-06-27")
        self.assertIsNone(rows[1].debit)
        self.assertEqual(rows[1].credit, 1200.00)

    def test_positioned_ocr_accepts_withdrawal_and_deposit_headers(self):
        words = [
            ("Date", 20, 100),
            ("Description", 150, 100),
            ("Withdrawals", 500, 100),
            ("Deposits", 620, 100),
            ("Balance", 750, 100),
            ("Jul", 20, 150),
            ("2,", 60, 150),
            ("2025", 95, 150),
            ("PAYMENT", 170, 150),
            ("25.00", 505, 150),
            ("975.00", 745, 150),
        ]
        data = {
            "text": [word for word, _, _ in words],
            "left": [x for _, x, _ in words],
            "top": [y for _, _, y in words],
            "width": [55] * len(words),
            "height": [20] * len(words),
            "conf": [95] * len(words),
        }

        rows = extractor.parse_scanned_debit_credit_ocr_data(data)

        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0].debit, 25.00)
        self.assertIsNone(rows[0].credit)


if __name__ == "__main__":
    unittest.main()
