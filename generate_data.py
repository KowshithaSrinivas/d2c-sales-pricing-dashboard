import random
from datetime import datetime, timedelta
import csv

random.seed(42)

# ---- Reference data ----
regions = ["DACH", "UK & Ireland", "France", "Benelux", "Nordics", "Southern Europe", "North America"]
product_lines = ["Smart Ovens", "Coffee Systems", "Cordless Vacuums", "Dishwashers", "Washing Machines", "Small Kitchen Appliances"]
lead_sources = ["Web-shop Organic", "Paid Search", "Social Media Ads", "Affiliate Partner", "Email Campaign", "Retail Referral", "Trade Show"]
deal_stages_won = "Closed Won"
deal_stages_lost = "Closed Lost"
sales_reps = ["A. Becker", "L. Dubois", "M. Hoffmann", "S. Andersson", "J. Murphy", "P. Rossi", "K. Nilsson", "T. Wagner"]

stage_funnel = ["New Lead", "Qualified", "Proposal Sent", "Negotiation", "Closed Won", "Closed Lost"]

base_price = {
    "Smart Ovens": 899,
    "Coffee Systems": 549,
    "Cordless Vacuums": 379,
    "Dishwashers": 749,
    "Washing Machines": 699,
    "Small Kitchen Appliances": 129,
}

start_date = datetime(2025, 1, 1)
end_date = datetime(2026, 6, 1)

def random_date(start, end):
    delta = end - start
    return start + timedelta(days=random.randint(0, delta.days))

rows = []
deal_id_counter = 10000

for i in range(1, 851):
    deal_id_counter += 1
    region = random.choice(regions)
    product = random.choice(product_lines)
    source = random.choices(
        lead_sources,
        weights=[25, 20, 18, 12, 10, 10, 5]
    )[0]
    rep = random.choice(sales_reps)

    created = random_date(start_date, end_date - timedelta(days=5))

    # Stage outcome distribution
    outcome_roll = random.random()
    if outcome_roll < 0.38:
        stage = "Closed Won"
    elif outcome_roll < 0.63:
        stage = "Closed Lost"
    elif outcome_roll < 0.75:
        stage = "Negotiation"
    elif outcome_roll < 0.88:
        stage = "Proposal Sent"
    else:
        stage = "Qualified"

    # Sales cycle length depends on stage
    if stage in ("Closed Won", "Closed Lost"):
        cycle_days = max(3, int(random.gauss(21, 9)))
        closed_date = created + timedelta(days=cycle_days)
        if closed_date > end_date:
            closed_date = end_date
    else:
        cycle_days = None
        closed_date = None

    unit_price = base_price[product]
    # add some price variance (promos, regional pricing)
    price_variance = random.uniform(0.85, 1.08)
    deal_value = round(unit_price * price_variance * random.choice([1, 1, 1, 2, 3]), 2)

    discount_pct = round(random.choice([0, 0, 0, 0.05, 0.1, 0.15, 0.2]), 2)

    rows.append({
        "Deal ID": f"D-{deal_id_counter}",
        "Created Date": created.strftime("%Y-%m-%d"),
        "Closed Date": closed_date.strftime("%Y-%m-%d") if closed_date else "",
        "Region": region,
        "Product Line": product,
        "Lead Source": source,
        "Sales Rep": rep,
        "Stage": stage,
        "Deal Value (EUR)": deal_value,
        "Discount %": discount_pct,
        "Sales Cycle (Days)": cycle_days if cycle_days else "",
    })

# Write CRM export CSV (source 1 - intentionally a bit messy: some blanks, inconsistent casing)
with open("data/crm_export_raw.csv", "w", newline="", encoding="utf-8") as f:
    writer = csv.DictWriter(f, fieldnames=rows[0].keys())
    writer.writeheader()
    for r in rows:
        # introduce light messiness for realism (a few rows with missing rep, inconsistent region casing)
        if random.random() < 0.03:
            r2 = dict(r)
            r2["Sales Rep"] = ""
            writer.writerow(r2)
        elif random.random() < 0.02:
            r2 = dict(r)
            r2["Region"] = r2["Region"].upper()
            writer.writerow(r2)
        else:
            writer.writerow(r)

print(f"Generated {len(rows)} deal records -> crm_export_raw.csv")

# ---- Competitor pricing benchmark dataset (source 2) ----
competitors = ["Bosch/Siemens (Own)", "Miele", "AEG/Electrolux", "Philips", "Beko", "Samsung", "LG"]

pricing_rows = []
for product in product_lines:
    own_price = base_price[product]
    for comp in competitors:
        if comp == "Bosch/Siemens (Own)":
            price = own_price
        else:
            variance = random.uniform(0.82, 1.25)
            price = round(own_price * variance, 2)
        pricing_rows.append({
            "Product Line": product,
            "Competitor": comp,
            "List Price (EUR)": price,
            "Online Rating (out of 5)": round(random.uniform(3.6, 4.8), 1),
            "Delivery Time (Days)": random.choice([2, 3, 4, 5, 7, 10]),
        })

with open("data/competitor_pricing_raw.csv", "w", newline="", encoding="utf-8") as f:
    writer = csv.DictWriter(f, fieldnames=pricing_rows[0].keys())
    writer.writeheader()
    writer.writerows(pricing_rows)

print(f"Generated {len(pricing_rows)} competitor pricing records -> competitor_pricing_raw.csv")
