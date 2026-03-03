# generate_dataset.py
# Run: python generate_dataset.py
# Output: ecommerce_data.csv  (12,000 rows)

import csv
import random
from datetime import date, timedelta

random.seed(42)

CATEGORIES = ["Electronics", "Apparel", "Home & Garden", "Beauty", "Sports"]
REGIONS    = ["North", "South", "East", "West"]
CAMPAIGNS  = ["Google_Search", "Meta_Ads", "Email", "Organic", "Influencer"]

def gen_date(start="2024-01-01", end="2024-12-31"):
    s = date.fromisoformat(start)
    e = date.fromisoformat(end)
    return s + timedelta(days=random.randint(0, (e - s).days))

rows        = []
cust_pool   = [f"CUST{str(i).zfill(5)}" for i in range(1, 3001)]
first_order = {}
order_id    = 10000

for _ in range(12000):
    cust  = random.choice(cust_pool)
    odate = gen_date()
    cat   = random.choice(CATEGORIES)
    reg   = random.choice(REGIONS)
    camp  = random.choice(CAMPAIGNS)
    qty   = random.randint(1, 5)

    base = random.uniform(40, 250)
    if odate.month >= 8:
        base *= random.uniform(0.60, 0.85)   # H2 revenue decline

    revenue     = round(base * qty, 2)
    discount    = round(revenue * random.uniform(0, 0.15), 2)
    net_revenue = round(revenue - discount, 2)
    cogs        = round(net_revenue * random.uniform(0.52, 0.65), 2)
    mkt_spend   = round(random.uniform(5, 120), 2)
    is_new      = 1 if cust not in first_order else 0
    first_order[cust] = odate
    returned    = 1 if (odate.month >= 8 and random.random() < 0.08) or random.random() < 0.03 else 0
    inventory   = random.randint(50, 500)
    visits      = random.randint(1, 8)

    rows.append([
        order_id, cust, odate.isoformat(), cat, reg, camp,
        qty, round(revenue, 2), round(discount, 2), net_revenue,
        cogs, mkt_spend, visits, is_new, returned, inventory
    ])
    order_id += 1

HEADER = [
    "order_id", "customer_id", "order_date", "category", "region", "campaign",
    "quantity", "gross_revenue", "discount", "net_revenue",
    "cogs", "marketing_spend", "website_visits", "is_new_customer",
    "is_returned", "inventory_units"
]

with open("ecommerce_data.csv", "w", newline="") as f:
    writer = csv.writer(f)
    writer.writerow(HEADER)
    writer.writerows(rows)

print(f"Generated {len(rows):,} rows -> ecommerce_data.csv")