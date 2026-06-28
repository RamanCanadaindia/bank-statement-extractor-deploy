import unittest

from mortgage_calculator import (
    calculate_maximum_mortgage,
    canadian_monthly_rate,
    present_value_of_payments,
)


class MortgageCalculatorTests(unittest.TestCase):
    def test_uses_stress_test_and_gds_expenses(self):
        result = calculate_maximum_mortgage(
            annual_household_income=120000,
            contract_rate_pct=4.5,
            amortization_years=25,
            annual_property_taxes=4800,
            monthly_heating=150,
            monthly_condo_fees=500,
        )

        self.assertAlmostEqual(result.maximum_housing_cost, 3900)
        self.assertAlmostEqual(result.maximum_mortgage_payment, 3100)
        self.assertAlmostEqual(result.qualifying_rate, 0.065)
        self.assertGreater(result.maximum_mortgage, 450000)
        self.assertLess(result.maximum_mortgage, 500000)

    def test_uses_floor_when_contract_plus_buffer_is_lower(self):
        result = calculate_maximum_mortgage(
            annual_household_income=80000,
            contract_rate_pct=2.0,
            amortization_years=25,
            annual_property_taxes=0,
            monthly_heating=0,
            monthly_condo_fees=0,
        )
        self.assertAlmostEqual(result.qualifying_rate, 0.0525)

    def test_no_capacity_when_fixed_costs_exceed_gds_limit(self):
        result = calculate_maximum_mortgage(
            annual_household_income=30000,
            contract_rate_pct=5.0,
            amortization_years=25,
            annual_property_taxes=12000,
            monthly_heating=500,
            monthly_condo_fees=1000,
        )
        self.assertEqual(result.maximum_mortgage_payment, 0)
        self.assertEqual(result.maximum_mortgage, 0)

    def test_zero_rate_present_value(self):
        self.assertEqual(present_value_of_payments(1000, 0, 12), 12000)
        self.assertEqual(canadian_monthly_rate(0), 0)


if __name__ == "__main__":
    unittest.main()
