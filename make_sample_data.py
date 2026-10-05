"""Generate messy sample files to demo sales_report.py.

Each file imitates a different source: its own column names, date format,
delimiter and currency formatting, plus the duplicates and broken rows that
real exports contain. The data is random and does not describe a real business.

Usage:
    python make_sample_data.py
"""

import csv
import random
from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook

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


def random_sales(count):
    """Return `count` sales, no two alike, so every duplicate in the files is deliberate."""
    seen, sales = set(), []
    while len(sales) < count:
        day = START + timedelta(days=random.randint(0, 180))
        product = random.choice(list(PRODUCTS))
        quantity = random.randint(1, 5)
        if (day, product, quantity) not in seen:
            seen.add((day, product, quantity))
            sales.append((day, product, quantity, PRODUCTS[product]))
    return sales


def write_csv(name, header, rows, delimiter=","):
    OUT.mkdir(exist_ok=True)
    with open(OUT / name, "w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle, delimiter=delimiter)
        writer.writerow(header)
        writer.writerows(rows)
    print(f"Wrote {name}: {len(rows)} rows")


# Source 1: a tidy point-of-sale export. ISO dates, clean columns.
rows = [[d.isoformat(), p, q, f"{price:.2f}", "NSW"] for d, p, q, price in random_sales(60)]
write_csv("store_sydney.csv", ["Date", "Product", "Qty", "Unit Price", "Region"], rows)

# Source 2: day-first dates, currency symbols, upper-case names, stray spaces.
# The export was run twice for a few days, so four rows appear twice.
rows = [
    [d.strftime("%d/%m/%Y"), p.upper(), q, f"${price:,.2f}", " vic "]
    for d, p, q, price in random_sales(50)
]
rows += rows[:4]
rows.append(["not a date", "Wireless Mouse", 2, "$29.90", "vic"])
rows.append(["12/03/2026", "USB-C Hub", "", "$59.50", "vic"])
write_csv("store_melbourne.csv", ["order_date", "item", "quantity", "price", "store"], rows)

# Source 3: a web shop export. Semicolon-delimited, text dates, lower-case names.
rows = [
    [d.strftime("%b %d, %Y"), p.lower(), q, f"{price:.2f}", "Online"]
    for d, p, q, price in random_sales(40)
]
rows.append(["Feb 30, 2026", "laptop stand", 1, "45.00", "Online"])  # this date does not exist
rows.append(["Mar 05, 2026", "", 3, "45.00", "Online"])
rows.append(["Apr 11, 2026", "webcam 1080p", -2, "89.00", "Online"])
write_csv(
    "online_orders.csv", ["Sold_On", "Product_Name", "Units", "Price_Each", "Location"], rows, delimiter=";"
)

# Source 4: an Excel workbook from the accounts system, with invoice numbers,
# real date cells and a TOTAL row at the bottom.
sales = random_sales(35)
rows = [[f"INV-{1001 + i}", d, p, q, price, "QLD"] for i, (d, p, q, price) in enumerate(sales)]
rows.append(list(rows[2]))   # invoice line exported twice: same invoice number
rows.append(list(rows[9]))
repeat = list(rows[5])       # a genuine repeat sale: same details, new invoice number
repeat[0] = "INV-1036"
rows.append(repeat)
total_units = sum(row[3] for row in rows)

workbook = Workbook()
sheet = workbook.active
sheet.title = "Invoices"
sheet.append(["Invoice No", "Invoice Date", "Description", "Units Sold", "Unit Price", "State"])
for row in rows:
    sheet.append(row)
sheet.append([None, None, "TOTAL", total_units, None, None])
OUT.mkdir(exist_ok=True)
workbook.save(OUT / "wholesale_brisbane.xlsx")
print(f"Wrote wholesale_brisbane.xlsx: {len(rows) + 1} rows")
