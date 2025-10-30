# test.py
import os
from generation import generate_dpr_package  # adjust if generation.py is elsewhere
from dotenv import load_dotenv

# Load environment variables (e.g., API keys)
load_dotenv()

# ✅ Example input data (same structure as ProjectRequest)
data = {
    "title": "Organic Agro Processing Unit",
    "short_description": "Establishing a small-scale agro processing unit for producing organic millet-based snacks and flours.",
    "location": "Anantapur, Andhra Pradesh",
    "sector": "agro_processing",
    "capacity": 500,  # in kilograms per day
    "currency": "INR",
    "estimated_cost": 3500000,  # in INR
    "promoter_name": "Sree Lakshmi Agro Foods",
    "contact_email": "contact@slagrofoods.in",
    "contact_phone": "+91-9876543210",
    "employment_generated": 25,
    "project_duration_months": 18,
    "objectives": [
        "Promote value addition in local millet production.",
        "Generate employment for rural youth and women.",
        "Create healthy, organic food products for urban markets."
    ],
    "funding_required": 2500000,
    "funding_type": "Term Loan",
    "land_area": "1.5 acres",
    "infrastructure_details": "Includes processing shed, storage unit, packaging area, and solar drying unit.",
    "technology_used": "Cold-press milling and automated packaging line with IoT sensors for moisture control.",
    "expected_revenue_year_1": 1200000,
    "expected_revenue_year_2": 1800000,
    "expected_revenue_year_3": 2500000
}


try:
    # Directly call the same function used by the API endpoint
    result = generate_dpr_package(data)
    print("✅ DPR Generation Successful!\n")
    print(result)
except Exception as e:
    print("❌ Error during DPR generation:")
    print(e)
