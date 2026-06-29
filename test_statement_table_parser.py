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


if __name__ == "__main__":
    unittest.main()
