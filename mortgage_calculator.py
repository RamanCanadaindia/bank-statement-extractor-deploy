from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class GDSResult:
    gross_monthly_income: float
    maximum_housing_cost: float
    property_tax_monthly: float
    heating_monthly: float
    condo_fee_portion: float
    maximum_mortgage_payment: float
    qualifying_rate: float
    maximum_mortgage: float
    estimated_purchase_price: float


def canadian_monthly_rate(annual_rate: float) -> float:
    """Convert a Canadian nominal semi-annual mortgage rate to a monthly rate."""
    if annual_rate < 0:
        raise ValueError("The interest rate cannot be negative.")
    return (1 + annual_rate / 2) ** (2 / 12) - 1


def present_value_of_payments(monthly_payment: float, annual_rate: float, months: int) -> float:
    if monthly_payment <= 0 or months <= 0:
        return 0.0
    monthly_rate = canadian_monthly_rate(annual_rate)
    if monthly_rate == 0:
        return monthly_payment * months
    return monthly_payment * (1 - (1 + monthly_rate) ** -months) / monthly_rate


def calculate_maximum_mortgage(
    *,
    annual_household_income: float,
    contract_rate_pct: float,
    amortization_years: int,
    annual_property_taxes: float,
    monthly_heating: float,
    monthly_condo_fees: float,
    gds_limit_pct: float = 39.0,
    stress_test_floor_pct: float = 5.25,
    stress_test_buffer_pct: float = 2.0,
    down_payment_pct: float = 20.0,
) -> GDSResult:
    """Calculate mortgage capacity using only the Gross Debt Service ratio."""
    numeric_values = {
        "Annual household income": annual_household_income,
        "Contract rate": contract_rate_pct,
        "Annual property taxes": annual_property_taxes,
        "Monthly heating": monthly_heating,
        "Monthly condo fees": monthly_condo_fees,
        "GDS limit": gds_limit_pct,
        "Stress-test floor": stress_test_floor_pct,
        "Stress-test buffer": stress_test_buffer_pct,
        "Down payment": down_payment_pct,
    }
    for label, value in numeric_values.items():
        if value < 0:
            raise ValueError(f"{label} cannot be negative.")
    if annual_household_income <= 0:
        raise ValueError("Annual household income must be greater than zero.")
    if amortization_years <= 0:
        raise ValueError("Amortization must be greater than zero.")
    if not 0 < gds_limit_pct <= 100:
        raise ValueError("GDS limit must be between 0 and 100 percent.")
    if not 0 <= down_payment_pct < 100:
        raise ValueError("Down payment must be less than 100 percent.")

    gross_monthly_income = annual_household_income / 12
    maximum_housing_cost = gross_monthly_income * gds_limit_pct / 100
    property_tax_monthly = annual_property_taxes / 12
    condo_fee_portion = monthly_condo_fees * 0.5
    maximum_payment = max(
        0.0,
        maximum_housing_cost - property_tax_monthly - monthly_heating - condo_fee_portion,
    )
    qualifying_rate = max(
        contract_rate_pct + stress_test_buffer_pct,
        stress_test_floor_pct,
    ) / 100
    maximum_mortgage = present_value_of_payments(
        maximum_payment,
        qualifying_rate,
        amortization_years * 12,
    )
    estimated_purchase_price = maximum_mortgage / (1 - down_payment_pct / 100)

    return GDSResult(
        gross_monthly_income=gross_monthly_income,
        maximum_housing_cost=maximum_housing_cost,
        property_tax_monthly=property_tax_monthly,
        heating_monthly=monthly_heating,
        condo_fee_portion=condo_fee_portion,
        maximum_mortgage_payment=maximum_payment,
        qualifying_rate=qualifying_rate,
        maximum_mortgage=maximum_mortgage,
        estimated_purchase_price=estimated_purchase_price,
    )
