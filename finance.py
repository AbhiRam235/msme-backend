# backend/finance.py
import pandas as pd
import numpy as np

def generate_financials(basic_inputs: dict):
    """
    Generate simple 5-year financial projection for DPR.
    Includes Revenue, Variable Cost, Fixed Cost, EBITDA.
    Returns: (DataFrame, meta dict)
    """
    capacity = basic_inputs.get("capacity") or 1000
    currency = basic_inputs.get("currency", "INR")
    year_count = 5

    price_per_unit = basic_inputs.get("additional", {}).get("price_per_unit", 100)
    variable_cost_per_unit = basic_inputs.get("additional", {}).get("variable_cost", 40)
    fixed_annual_overheads = basic_inputs.get("additional", {}).get("fixed_annual", 200000)

    years = [f"Year {i+1}" for i in range(year_count)]
    revenue = [capacity * price_per_unit * (1 + 0.05 * i) for i in range(year_count)]
    variable_cost = [capacity * variable_cost_per_unit * (1 + 0.03 * i) for i in range(year_count)]
    fixed_cost = [fixed_annual_overheads * (1 + 0.04 * i) for i in range(year_count)]
    ebitda = [r - vc - fc for r, vc, fc in zip(revenue, variable_cost, fixed_cost)]

    df = pd.DataFrame({
        "Year": years,
        "Revenue": revenue,
        "Variable Cost": variable_cost,
        "Fixed Cost": fixed_cost,
        "EBITDA": ebitda
    })
    df.set_index("Year", inplace=True)

    capex = basic_inputs.get("additional", {}).get("capex", 5_000_000)
    opex_breakdown = {
        "labor": basic_inputs.get("additional", {}).get("labor", 500000),
        "maintenance": basic_inputs.get("additional", {}).get("maintenance", 200000),
        "utilities": basic_inputs.get("additional", {}).get("utilities", 150000)
    }

    meta = {"capex": capex, "opex_breakdown": opex_breakdown, "currency": currency}
    return df, meta
