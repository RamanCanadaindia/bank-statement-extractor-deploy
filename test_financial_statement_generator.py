from datetime import date
import io
import unittest
from unittest.mock import patch

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

    def test_separate_description_code_and_amount_lines_are_supported(self):
        result = generator.parse_schedule_text(
            "Cash\n1001\n25,000\nTotal assets\n2599\n25,000",
            "100",
        )
        self.assertEqual(result.entries.iloc[0]["Description"], "Cash")
        self.assertEqual(float(result.entries.iloc[0]["Amount"]), 25000)

    def test_known_labels_recover_rows_when_ocr_misses_gifi_codes(self):
        result = generator.parse_schedule_text(
            "Cash $25,000\n"
            "Total assets $25,000\n"
            "Total liabilities $10,000\n"
            "Total shareholder equity $15,000\n"
            "Total liabilities and shareholder equity $25,000",
            "100",
        )
        self.assertEqual(
            float(result.entries.loc[result.entries["Code"] == 2599, "Amount"].iloc[0]),
            25000,
        )
        self.assertEqual(result.warnings, [])

    def test_acroform_gifi_fields_are_supported(self):
        from reportlab.pdfgen import canvas

        buffer = io.BytesIO()
        pdf = canvas.Canvas(buffer)
        fields = {
            "schedule100_gifi_1001": "25000",
            "schedule100_gifi_2599": "25000",
            "schedule100_gifi_3499": "10000",
            "schedule100_gifi_3620": "15000",
            "schedule100_gifi_3640": "25000",
        }
        y = 740
        for name, value in fields.items():
            pdf.acroForm.textfield(name=name, value=value, x=72, y=y, width=120, height=16)
            y -= 24
        pdf.showPage()
        pdf.save()

        result = generator.extract_schedule_pdf(buffer.getvalue(), "100")
        self.assertIn(2599, result.entries["Code"].tolist())
        self.assertEqual(
            float(result.entries.loc[result.entries["Code"] == 2599, "Amount"].iloc[0]),
            25000,
        )

    @patch.object(generator, "extract_tesseract_text", return_value=SCHEDULE_100)
    @patch.object(generator, "extract_docling_text", return_value="")
    @patch.object(generator, "extract_pdf_form_text", return_value="")
    @patch.object(generator, "extract_pdf_layout_text", return_value="")
    @patch.object(generator, "extract_pdf_text", return_value="")
    def test_image_only_pdf_uses_tesseract_fallback(
        self,
        _pdf_text,
        _layout_text,
        _form_text,
        _docling_text,
        tesseract_text,
    ):
        result = generator.extract_schedule_pdf(b"image-only-pdf", "100")
        self.assertIn(2599, result.entries["Code"].tolist())
        tesseract_text.assert_called_once()

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

    def test_condensed_gifi_uses_current_year_not_prior_year(self):
        text = """Name of corporation : 1386371 B.C. LTD. 2025/06/30
Deposits in Canadian banks and institutions: CDN currency 1002 + 7,812 2,089
Total current assets 1599 = 7,812 2,089
Motor vehicles 1742 + 16,500
Accumulated amortization of motor vehicles 1743 - 2,475
Total tangible capital assets 2008 = 16,500
Total accumulated amortization of capital assets 2009 = 2,475
Total Assets 2599 = 21,837 2,089
Taxes payable 2680 + 907 215
Total current liabilities 3139 = 907 215
Due to individual shareholder(s) 3261 + 17,253 2,113
Common shares 3500 + 100 100
Retained earnings/deficit 3600 + 3,577 (339)
Total shareholder equity 3620 = 3,677 (239)
Total Liabilities and Shareholder's Equity = 21,837 2,089
"""
        result = generator.parse_schedule_text(text, "100")
        self.assertEqual(float(result.entries.loc[result.entries["Code"] == 1002, "Amount"].iloc[0]), 7812)
        self.assertEqual(float(result.entries.loc[result.entries["Code"] == 2680, "Amount"].iloc[0]), 907)
        self.assertEqual(float(result.entries.loc[result.entries["Code"] == 1743, "Amount"].iloc[0]), -2475)
        self.assertEqual(float(result.entries.loc[result.entries["Code"] == 3499, "Amount"].iloc[0]), 18160)
        self.assertEqual(result.warnings, [])

    def test_combined_pdf_metadata_is_detected(self):
        from reportlab.pdfgen import canvas

        buffer = io.BytesIO()
        pdf = canvas.Canvas(buffer)
        pdf.drawString(72, 760, "Business Number : 791178619 Tax year end")
        pdf.drawString(72, 740, "Name of corporation : 1386371 B.C. LTD. 2025/06/30")
        pdf.drawString(72, 720, "BN: 791178619RC0001")
        pdf.save()
        metadata = generator.extract_statement_metadata(buffer.getvalue())
        self.assertEqual(metadata["company_name"], "1386371 B.C. LTD.")
        self.assertEqual(metadata["business_number"], "791178619RC0001")
        self.assertEqual(metadata["year_end"], date(2025, 6, 30))

    def test_report_outputs_are_created(self):
        pdf = generator.build_financial_statement_pdf(self.balance.entries, self.income.entries, self.metadata)
        docx = generator.build_financial_statement_docx(self.balance.entries, self.income.entries, self.metadata)
        self.assertTrue(pdf.startswith(b"%PDF"))
        self.assertTrue(docx.startswith(b"PK"))
        self.assertGreater(len(pdf), 5000)
        self.assertGreater(len(docx), 5000)


if __name__ == "__main__":
    unittest.main()
