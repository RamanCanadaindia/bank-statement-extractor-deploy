from datetime import date
import io
import unittest

import financial_statement_generator as generator


SCHEDULE_100 = """Schedule 100 Balance Sheet Information
1001 Cash 25,000
1060 Accounts receivable 35,000
1740 Machinery and equipment 50,000
1787 Accumulated amortization of machinery and equipment (10,000)
2599 Total assets 100,000
2621 Trade payables 20,000
3143 Chartered bank loan 40,000
3499 Total liabilities 60,000
3500 Common shares 1,000
3600 Retained earnings 39,000
3620 Total shareholder equity 40,000
3640 Total liabilities and shareholder equity 100,000
"""

SCHEDULE_125 = """Schedule 125 Income Statement Information
8000 Trade sales of goods and services 150,000
8299 Total revenue 150,000
8320 Purchases 40,000
8518 Cost of sales 40,000
8519 Gross profit 110,000
8521 Advertising 5,000
8715 Bank charges 1,000
8810 Office expenses 4,000
8862 Accounting fees 3,000
9060 Salaries and wages 60,000
9220 Utilities 7,000
9367 Total operating expenses 80,000
9368 Total expenses 120,000
9970 Net income before taxes 30,000
9990 Current income taxes 3,500
9999 Net income after taxes 26,500
"""


class FinancialStatementGeneratorTests(unittest.TestCase):
    def setUp(self):
        self.balance = generator.parse_schedule_text(SCHEDULE_100, "100")
        self.income = generator.parse_schedule_text(SCHEDULE_125, "125")
        self.metadata = {
            "company_name": "Example Transport Ltd.",
            "business_number": "123456789RC0001",
            "year_end": date(2025, 12, 31),
            "firm_name": "Raman Financial Services",
            "firm_address": "Surrey, British Columbia",
            "report_date": date(2026, 3, 15),
            "basis_of_accounting": "Accrual basis selected by management.",
            "report_text": generator.DEFAULT_COMPILATION_REPORT,
        }

    def test_balanced_schedules_have_no_warnings(self):
        self.assertEqual(self.balance.warnings, [])
        self.assertEqual(self.income.warnings, [])
        self.assertEqual(float(self.balance.entries.loc[self.balance.entries["Code"] == 2599, "Amount"].iloc[0]), 100000)
        self.assertEqual(float(self.income.entries.loc[self.income.entries["Code"] == 9999, "Amount"].iloc[0]), 26500)

    def test_code_amount_layout_uses_preceding_description(self):
        result = generator.parse_schedule_text("Cash\n1001 25,000\nTotal assets\n2599 25,000", "100")
        self.assertEqual(result.entries.iloc[0]["Description"], "Cash")

    def test_unbalanced_schedule_is_flagged(self):
        result = generator.parse_schedule_text(SCHEDULE_100.replace("3640 Total liabilities and shareholder equity 100,000", "3640 Total liabilities and shareholder equity 99,000"), "100")
        self.assertTrue(any("does not agree" in warning for warning in result.warnings))

    def test_searchable_pdf_upload_is_extracted(self):
        from reportlab.pdfgen import canvas

        buffer = io.BytesIO()
        pdf = canvas.Canvas(buffer)
        y = 760
        for line in SCHEDULE_100.splitlines():
            pdf.drawString(72, y, line)
            y -= 15
        pdf.save()
        result = generator.extract_schedule_pdf(buffer.getvalue(), "100")
        self.assertEqual(result.warnings, [])
        self.assertIn(2599, result.entries["Code"].tolist())

    def test_report_outputs_are_created(self):
        pdf = generator.build_financial_statement_pdf(self.balance.entries, self.income.entries, self.metadata)
        docx = generator.build_financial_statement_docx(self.balance.entries, self.income.entries, self.metadata)
        self.assertTrue(pdf.startswith(b"%PDF"))
        self.assertTrue(docx.startswith(b"PK"))
        self.assertGreater(len(pdf), 5000)
        self.assertGreater(len(docx), 5000)


if __name__ == "__main__":
    unittest.main()
