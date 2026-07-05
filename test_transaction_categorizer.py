import unittest

import pandas as pd

import transaction_categorizer as categorizer


class TransactionCategorizerTests(unittest.TestCase):
    def test_direction_sensitive_rules(self):
        category, confidence, _ = categorizer.categorize_description(
            "STRIPE PAYOUT", 1200.00
        )
        self.assertEqual(category, "Sales Revenue")
        self.assertEqual(confidence, "High")

    def test_uber_is_transportation(self):
        category, _, _ = categorizer.categorize_description("UBER TRIP", -24.50)
        self.assertEqual(category, "Travel/Transportation")

    def test_custom_rules_take_priority(self):
        custom = [
            categorizer.CategoryRule(("amazon",), "Computer Equipment", "Debit", "Custom")
        ]
        category, confidence, _ = categorizer.categorize_description(
            "AMAZON MARKETPLACE", -850.00, custom
        )
        self.assertEqual(category, "Computer Equipment")
        self.assertEqual(confidence, "Custom")

    def test_existing_categories_and_balance_rows_are_preserved(self):
        frame = pd.DataFrame(
            [
                {"Description": "Opening Balance", "Amount": 1000, "Category": "Opening Balance"},
                {"Description": "SHELL", "Amount": -50, "Category": "Vehicle Expenses"},
            ]
        )
        result = categorizer.categorize_transactions(frame)
        self.assertEqual(result.iloc[0]["Category"], "Opening Balance")
        self.assertEqual(result.iloc[1]["Category"], "Vehicle Expenses")

    def test_debit_credit_columns_become_signed_amount(self):
        frame = pd.DataFrame(
            [
                {"Description": "FEE", "Debit": 12.0, "Credit": None},
                {"Description": "DEPOSIT", "Debit": None, "Credit": 500.0},
            ]
        )
        payload = frame.to_csv(index=False).encode()
        loaded = categorizer.load_transactions(payload, "transactions.csv")
        self.assertEqual(loaded["Amount"].tolist(), [-12.0, 500.0])

    def test_custom_rule_template_is_valid_csv(self):
        template = pd.read_csv(
            __import__("io").BytesIO(categorizer.custom_rule_template())
        )
        self.assertEqual(
            template.columns.tolist(),
            ["Keyword", "Category", "Applies To"],
        )


if __name__ == "__main__":
    unittest.main()
