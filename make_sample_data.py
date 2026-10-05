"""Generate messy sample CSV files to demo sales_report.py.

Each file imitates a different source: different column names, date formats,
currency formatting, plus duplicates and broken rows.

Usage:
    python make_sample_data.py
"""

import csv
import random
from datetime import date, timedelta
from pathlib import Path

random.seed(42)

PRODUCTS = {
    "Wireless Mouse": 29.90,
    "Mechanical Keyboard": 119.00,
    "USB-C Hub": 59.50,
    "Laptop Stand": 45.00,
    "Webcam 1080p": 89.00,
    "Monitor 27 inch": 349.00,
}
START = date(2026, 1, 1)
OUT = Path("data")


def random_sale():
    day = START + timedelta(days=random.randint(0, 180))
    product = random.choice(list(PRODUCTS))
    return day, product, random.randint(1, 5), PRODUCTS[product]


def write(name, header, rows):
    OUT.mkdir(exist_ok=True)
    with open(OUT / name, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(header)
        writer.writerows(rows)
    print(f"Wrote {name}: {len(rows)} rows")


# Source 1: clean ISO dates, tidy columns.
rows = []
for _ in range(60):
    d, p, q, price = random_sale()
    rows.append([d.isoformat(), p, q, f"{price:.2f}", "NSW"])
write("store_sydney.csv", ["Date", "Product", "Qty", "Unit Price", "Region"], rows)

# Source 2: Australian dates, currency symbols, inconsistent casing, duplicates.
rows = []
for _ in range(50):
    d, p, q, price = random_sale()
    rows.append([d.strftime("%d/%m/%Y"), p.upper(), q, f"${price:,.2f}", " vic "])
rows += rows[:4]  # exact duplicates
rows.append(["not a date", "Wireless Mouse", 2, "$29.90", "vic"])
rows.append(["12/03/2026", "USB-C Hub", "", "$59.50", "vic"])
write("store_melbourne.csv", ["order_date", "item", "quantity", "price", "store"], rows)

# Source 3: text dates, different names again, negative and missing values.
rows = []
for _ in range(40):
    d, p, q, price = random_sale()
    rows.append([d.strftime("%b %d, %Y"), p.lower(), q, f"{price:.2f}", "Online"])
rows.append(["Feb 30, 2026", "laptop stand", 1, "45.00", "Online"])
rows.append(["Mar 05, 2026", "", 3, "45.00", "Online"])
rows.append(["Apr 11, 2026", "webcam 1080p", -2, "89.00", "Online"])
write("online_orders.csv", ["Sold_On", "Product_Name", "Units", "Price_Each", "Location"], rows)
