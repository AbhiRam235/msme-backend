# backend/templates.py
# Carefully crafted DPR templates for common MSME project classes.
# Each template lists the recommended top-level sections for a bankable DPR.
# Use these as the basis for generation (models can expand each section into subsections).

PROJECT_TEMPLATES = {
    "agro_processing": {
        "name": "Agro Processing Plant",
        "sections": [
            "Executive Summary",
            "Project Background & Objectives",
            "Product & Process Description",
            "Raw Material & Supply Chain",
            "Technical Feasibility & Plant Layout",
            "Location & Site Assessment",
            "Market Analysis & Demand Forecast",
            "Sales & Marketing Strategy",
            "Operations Plan (Capacity, Manpower)",
            "Quality Control & Certifications",
            "Environmental & Regulatory Compliance",
            "Implementation Schedule & Milestones",
            "Financial Projections (5-year)",
            "Funding Requirement & Utilization of Funds",
            "Risk Assessment & Mitigation Measures",
            "Monitoring, KPIs & Exit Strategy",
            "Annexures (BOQ, Vendor Quotations, Permits)"
        ]
    },

    "textile_manufacturing": {
        "name": "Textile Manufacturing Unit",
        "sections": [
            "Executive Summary",
            "Business Rationale & Background",
            "Products & Specifications (Yarn/Fabric/Garments)",
            "Technology, Machinery & Plant Layout",
            "Raw Materials & Sourcing Strategy",
            "Production Process & Capacity Planning",
            "Quality Assurance & Compliance (Standards)",
            "Market Analysis (Domestic & Export)",
            "Sales Channels & Pricing Strategy",
            "Operations & HR Plan",
            "Logistics, Warehousing & Inventory Management",
            "Environmental Considerations (effluent, waste)",
            "Implementation Timeline & Milestones",
            "Financial Projections (Costs, Revenue, EBITDA)",
            "Funding & Financial Structure",
            "Risk Analysis & Contingency Plans",
            "Annexures (Technical Specs, Vendor Quotes)"
        ]
    },

    "solar_energy": {
        "name": "Solar Power Project (Rooftop / Ground Mounted)",
        "sections": [
            "Executive Summary",
            "Project Overview & Objectives",
            "Site Selection & Solar Resource Assessment",
            "System Design & Technology (Panels, Inverters, Mounting)",
            "Capacity & Energy Yield Estimates",
            "Grid Connectivity / Off-grid Considerations",
            "Procurement, Vendors & EPC Plan",
            "Operation & Maintenance (O&M) Plan",
            "Financial Model (Tariff, PPA assumptions, ROI)",
            "Incentives, Subsidies & Policy Support",
            "Environmental & Social Impact Assessment",
            "Permits, Approvals & Land/Lease Details",
            "Implementation Schedule",
            "Risk Management (technical, policy, weather)",
            "Monitoring, Performance Guarantees & KPIs",
            "Annexures (Irradiance Data, Vendor Quotes)"
        ]
    },

    "wind_energy": {
        "name": "Wind Energy Project",
        "sections": [
            "Executive Summary",
            "Project Rationale & Site Overview",
            "Wind Resource Assessment & Met Mast Data",
            "Turbine Technology & Layout",
            "Grid Connection & Power Evacuation",
            "Civil Works & Foundation Requirements",
            "Procurement & EPC Strategy",
            "Operations, Maintenance & Spares",
            "Financial Viability (LCOE, Cash Flows)",
            "Permits, Environmental Clearances & Land Use",
            "Community Engagement & Social Safeguards",
            "Implementation Timeline",
            "Risk & Mitigation (wind variability, grid risk)",
            "Monitoring & Performance Reporting",
            "Annexures (Wind Roses, Vendor Specs)"
        ]
    },

    "food_packaging": {
        "name": "Food Packaging & Processing",
        "sections": [
            "Executive Summary",
            "Background & Product Range",
            "Processing & Packaging Technology",
            "Raw Materials & Supplier Assessment",
            "Quality, Food Safety & Certifications (FSSAI, ISO)",
            "Production Capacity & Workflow",
            "Packaging Design & Shelf-life Considerations",
            "Market Research & Customer Segments",
            "Distribution & Cold Chain (if applicable)",
            "Operations & HR Plan",
            "Financial Projections & Break-even Analysis",
            "Regulatory Compliance & Labelling",
            "Implementation Plan & Timeline",
            "Risk Assessment (food safety, supply chain)",
            "Annexures (lab test reports, vendor quotes)"
        ]
    },

    "waste_management": {
        "name": "Waste Management & Recycling",
        "sections": [
            "Executive Summary",
            "Problem Statement & Project Rationale",
            "Types of Waste Covered (MSW, E-waste, Bio-waste)",
            "Technology & Process Flow (segregation, recycling, composting)",
            "Feedstock Supply & Collection Mechanism",
            "Site Requirements & Environmental Controls",
            "By-products & Revenue Streams (recyclates, compost, energy)",
            "Operations & Logistics (collection, transport, storage)",
            "Permits, Licenses & Environmental Clearances",
            "Financial Model & Funding Requirements",
            "Social Impact & Community Engagement",
            "Implementation Roadmap",
            "Risk & Mitigation (policy, contamination, market for recyclates)",
            "Monitoring & Impact Metrics",
            "Annexures (MSW studies, vendor proposals)"
        ]
    },

    "electric_vehicle_charging": {
        "name": "Electric Vehicle Charging Station",
        "sections": [
            "Executive Summary",
            "Project Background & Need Assessment",
            "Site Selection & Footfall Analysis",
            "Charging Technology & Equipment (AC/DC, CCS/CHAdeMO)",
            "Grid/Power Supply & Transformer/Panel Details",
            "Business Model (pay-per-use, subscription, fleet charging)",
            "Revenue Streams & Pricing Strategy",
            "Operations, Staffing & Safety Protocols",
            "Permits, Utility Agreements & Interconnection",
            "Financial Projections (uptake, payback period)",
            "Environmental Benefits & Electrification Impact",
            "Implementation Timeline & CAPEX Plan",
            "Risk & Contingency (power outages, low utilisation)",
            "Annexures (site maps, equipment quotes)"
        ]
    },

    "battery_manufacturing": {
        "name": "Battery Manufacturing / Cell Assembly",
        "sections": [
            "Executive Summary",
            "Market Opportunity & Strategic Fit",
            "Product Types (cells, packs, chemistry)",
            "Technology & Production Process",
            "Raw Materials, Sourcing & Supply Chain",
            "Plant Layout, Utilities & Safety (fire, handling)",
            "Quality Control, Testing & Standards",
            "Environmental & Hazardous Waste Management",
            "Workforce & Skill Requirements",
            "Financial Projections & Capital Structure",
            "Regulatory Compliance & Certifications",
            "Implementation Phases & Timelines",
            "Risk & Mitigation (raw material volatility, safety)",
            "Annexures (technical specs, vendor & capex quotes)"
        ]
    },

    "pharmaceuticals": {
        "name": "Pharmaceuticals / Formulation Unit",
        "sections": [
            "Executive Summary",
            "Business Case & Market Opportunity",
            "Product Portfolio & Dosage Forms",
            "Manufacturing Process & Technology",
            "Quality Assurance & Regulatory (GMP, CDSCO)",
            "Raw Materials & Vendor Qualification",
            "Plant Layout, Utilities & Clean Rooms",
            "Packaging, Labelling & Serialization",
            "Distribution, Cold Chain & Market Access",
            "Financial Model & Investment Plan",
            "Regulatory Approvals & Compliance Roadmap",
            "Implementation Schedule",
            "Risk Assessment (regulatory delays, quality issues)",
            "Annexures (spec sheets, validation plans)"
        ]
    },

    "healthcare_clinic": {
        "name": "Primary Healthcare Clinic / Diagnostic Centre",
        "sections": [
            "Executive Summary",
            "Community Health Needs & Location Analysis",
            "Services Offered (OPD, diagnostics, minor procedures)",
            "Facility Layout & Equipment List",
            "Staffing & HR Plan (doctors, nurses, technicians)",
            "Clinical Protocols & Quality Assurance",
            "Patient Flow, Records & Digital Health Integration",
            "Pricing & Revenue Model (consultations, tests, packages)",
            "Regulatory & Licenses (clinical, waste disposal)",
            "Financial Projections & Sustainability",
            "Partnerships with Insurers / Govt Schemes",
            "Implementation Timeline",
            "Risk & Mitigation (clinical risk, footfall shortfall)",
            "Annexures (equipment quotes, staffing rosters)"
        ]
    },

    "it_services": {
        "name": "IT & Software Services",
        "sections": [
            "Executive Summary",
            "Company Background & Services Portfolio",
            "Target Markets & Client Segments",
            "Technology Stack & IP Considerations",
            "Service Delivery Model (onshore/offshore/hybrid)",
            "Team Structure & Talent Acquisition",
            "Sales, Marketing & Client Acquisition Strategy",
            "Infrastructure, Security & Compliance",
            "Pricing, Contracts & SLA Model",
            "Implementation Roadmap & Milestones",
            "Financial Projections & Revenue Recognition",
            "Risk & Mitigation (talent churn, project delivery)",
            "KPIs & Monitoring",
            "Annexures (sample SOW, client references)"
        ]
    },

    "construction_materials": {
        "name": "Construction Materials Manufacturing",
        "sections": [
            "Executive Summary",
            "Market Demand & Usage (regional construction demand)",
            "Product Range (cement, bricks, tiles, aggregates)",
            "Raw Materials & Quarry/Source Details",
            "Manufacturing Process & Technology",
            "Plant Layout, Utilities & Emissions Control",
            "Quality Standards & Testing Protocols",
            "Logistics, Supply Chain & Distribution",
            "Sales Channels & Pricing Strategy",
            "Financial Projections & CAPEX Plan",
            "Environmental & Regulatory Compliance",
            "Implementation Timeline",
            "Risk Assessment (raw material shortage, market cycles)",
            "Annexures (technical drawings, vendor quotes)"
        ]
    },

    "hospitality_resort": {
        "name": "Hospitality / Resort Project",
        "sections": [
            "Executive Summary",
            "Concept & USP (target guests, theme)",
            "Site & Location Analysis",
            "Architectural Plan & Facilities",
            "Service Offerings (F&B, events, wellness)",
            "Operations & Staffing Plan",
            "Sales & Marketing Strategy (OTA, direct bookings)",
            "Revenue & Occupancy Projections",
            "Permits, Safety & Environmental Considerations",
            "Financial Model & Funding Requirement",
            "Implementation Timeline & Phased Opening",
            "Risk & Contingency (seasonality, demand shocks)",
            "Sustainability & Local Community Engagement",
            "Annexures (site plans, vendor quotes)"
        ]
    },

    "education_institute": {
        "name": "Education / Training Institute",
        "sections": [
            "Executive Summary",
            "Vision, Mission & Educational Objectives",
            "Courses & Curriculum Offered",
            "Target Student Segments & Market Research",
            "Campus / Infrastructure Requirements",
            "Faculty & Staff Recruitment Plan",
            "Pedagogy, Assessment & Accreditation",
            "Student Acquisition & Marketing",
            "Fee Structure & Revenue Model",
            "Operational Plan & Academic Calendar",
            "Financial Projections & Affordability Analysis",
            "Regulatory Compliance & Approvals",
            "Implementation Roadmap",
            "Risk Management (enrolment shortfall, faculty gaps)",
            "Annexures (course syllabi, faculty CVs)"
        ]
    },

    # --- Newly Added Templates ---
    "cloud_kitchen": {
        "name": "Cloud Kitchen",
        "sections": [
            "Executive Summary",
            "Business Model & Concept",
            "Menu Planning & Cuisine Strategy",
            "Kitchen Layout & Equipment List",
            "Raw Materials & Supplier Network",
            "Food Safety, Hygiene & Licensing (FSSAI)",
            "Technology Integration (POS, Delivery Platforms)",
            "Marketing & Branding Strategy",
            "Operations & Workforce Plan",
            "Financial Projections & Break-even Analysis",
            "Risk & Mitigation (food quality, delivery delays)",
            "Implementation Timeline",
            "Annexures (vendor quotes, menu samples)"
        ]
    },

    "boutique": {
        "name": "Fashion Boutique",
        "sections": [
            "Executive Summary",
            "Concept & Target Market",
            "Product Range (apparel, accessories)",
            "Store Layout & Visual Merchandising",
            "Sourcing & Supplier Strategy",
            "Marketing & Branding (social media, influencer tie-ups)",
            "Pricing & Sales Strategy",
            "Staffing & Customer Service Plan",
            "Financial Projections & Profitability",
            "Implementation Timeline",
            "Risk Assessment (inventory, fashion trends)",
            "Annexures (vendor lists, store design)"
        ]
    },

    "real_estate": {
        "name": "Real Estate Development Project",
        "sections": [
            "Executive Summary",
            "Project Concept & Location Details",
            "Market Analysis & Demand Assessment",
            "Architectural Design & Layout",
            "Construction Plan & Technology",
            "Regulatory Approvals & RERA Compliance",
            "Sales & Marketing Strategy",
            "Financial Model (ROI, Cash Flows)",
            "Implementation Schedule",
            "Risk Assessment (delays, legal, cost overruns)",
            "Sustainability Features",
            "Annexures (site plan, approvals)"
        ]
    },

    "photo_studio": {
        "name": "Photography & Videography Studio",
        "sections": [
            "Executive Summary",
            "Services Offered (studio shoots, events, editing)",
            "Equipment & Infrastructure",
            "Team Structure & Roles",
            "Marketing & Client Acquisition Strategy",
            "Pricing & Packages",
            "Financial Projections & ROI",
            "Operational Plan",
            "Risk & Contingency (equipment downtime, low bookings)",
            "Annexures (sample portfolio, vendor quotes)"
        ]
    },

    "bakery": {
        "name": "Bakery & Confectionery Unit",
        "sections": [
            "Executive Summary",
            "Product Range (breads, cakes, pastries)",
            "Production Process & Equipment",
            "Raw Material Sourcing",
            "Food Safety & Hygiene Standards",
            "Packaging & Branding",
            "Marketing & Distribution Strategy",
            "Financial Projections & Break-even Point",
            "Implementation Timeline",
            "Risk & Mitigation (perishables, quality issues)",
            "Annexures (recipes, vendor details)"
        ]
    },

    "supermarket": {
        "name": "Supermarket / Retail Store",
        "sections": [
            "Executive Summary",
            "Concept & Product Mix",
            "Store Layout & Design",
            "Supply Chain & Vendor Management",
            "Pricing & Inventory Strategy",
            "Staffing & Operations Plan",
            "Marketing & Customer Loyalty Programs",
            "Financial Projections (Revenue, ROI)",
            "Implementation Schedule",
            "Risk & Mitigation (stockouts, competition)",
            "Annexures (vendor contracts, layout plans)"
        ]
    },

    "ice_cream": {
        "name": "Ice Cream Factory / Shop",
        "sections": [
            "Executive Summary",
            "Product Range & Flavour Development",
            "Production Technology & Machinery",
            "Raw Materials & Supply Chain",
            "Quality Control & Food Safety (FSSAI)",
            "Branding, Packaging & Distribution",
            "Sales & Marketing Strategy",
            "Financial Projections & Payback Analysis",
            "Implementation Timeline",
            "Risk & Mitigation (storage, seasonal demand)",
            "Annexures (equipment quotes, packaging samples)"
        ]
    },

    "organic_farm": {
        "name": "Organic Farming Project",
        "sections": [
            "Executive Summary",
            "Project Objectives & Location Details",
            "Crop Selection & Cultivation Plan",
            "Organic Inputs & Certification Process",
            "Irrigation & Farm Infrastructure",
            "Harvesting, Storage & Value Addition",
            "Market Analysis & Distribution Channels",
            "Financial Projections & Subsidy Support",
            "Risk & Mitigation (pest, climate, certification)",
            "Implementation Timeline",
            "Annexures (land layout, input supplier details)"
        ]
    },

    "car_garage": {
        "name": "Car Service Garage / Auto Workshop",
        "sections": [
            "Executive Summary",
            "Service Offerings (repairs, detailing, diagnostics)",
            "Equipment & Tools Required",
            "Manpower & Skills Plan",
            "Customer Segments & Pricing",
            "Marketing & Brand Positioning",
            "Operations & Workflow",
            "Financial Projections & Break-even Analysis",
            "Risk & Mitigation (accidents, low footfall)",
            "Implementation Timeline",
            "Annexures (vendor list, layout design)"
        ]
    },

    "fitness_wellness": {
        "name": "Fitness & Wellness Center",
        "sections": [
            "Executive Summary",
            "Concept & Service Offerings (gym, yoga, spa)",
            "Facility Layout & Equipment",
            "Trainer & Staff Requirements",
            "Membership Pricing & Packages",
            "Marketing & Retention Strategy",
            "Financial Projections & ROI",
            "Implementation Timeline",
            "Risk & Mitigation (injuries, churn)",
            "Annexures (vendor quotes, trainer profiles)"
        ]
    },

    "sports_facility": {
        "name": "Sports Facility (Badminton / Tennis / Box Cricket / Table Tennis)",
        "sections": [
            "Executive Summary",
            "Project Concept & Location",
            "Sports Infrastructure & Design",
            "Equipment & Amenities",
            "Membership / Booking Model",
            "Marketing & Sponsorship Strategy",
            "Operations & Maintenance Plan",
            "Financial Projections & ROI",
            "Risk & Mitigation (injury, weather, utilization)",
            "Implementation Timeline",
            "Annexures (layout drawings, vendor quotes)"
        ]
    },

    # fallback default template
    "default": {
        "name": "General Project",
        "sections": [
            "Executive Summary",
            "Project Description",
            "Market Analysis",
            "Technical & Operational Plan",
            "Implementation Schedule",
            "Financial Projections",
            "Funding Requirement",
            "Risk & Mitigation",
            "Monitoring & KPIs",
            "Annexures"
        ]
    }
}

