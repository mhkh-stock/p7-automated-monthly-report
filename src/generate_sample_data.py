"""Generate SAMPLE monthly sales data for CityMart (fictional shop).
Deterministic (seed 42): re-running reproduces the identical CSV.
Months: Jul, Aug, Sep 2026. One row per order line."""
import csv, random
from pathlib import Path
random.seed(42)
OUT = Path(__file__).resolve().parents[1] / "data" / "raw" / "sales_SAMPLE.csv"
PRODUCTS = [
    ("LED Desk Lamp", "Electronics", 4200), ("Bluetooth Speaker", "Electronics", 6500),
    ("Wireless Mouse", "Electronics", 2800), ("Stainless Steel Kettle", "Kitchen", 5400),
    ("Non-stick Pan Set", "Kitchen", 8900), ("Ceramic Dinner Set", "Kitchen", 12500),
    ("Office Chair", "Furniture", 18900), ("Wooden Bookshelf", "Furniture", 15400),
    ("Notebook Pack (5)", "Stationery", 1200), ("Pen Set (10)", "Stationery", 950),
]
rows = []
oid = 1000
for ym, n in [("2026-07", 170), ("2026-08", 185), ("2026-09", 205)]:
    y, m = map(int, ym.split("-"))
    for _ in range(n):
        day = random.randint(1, 28)
        name, cat, price = random.choice(PRODUCTS)
        qty = random.choice([1, 1, 1, 2, 2, 3])
        oid += 1
        rows.append([f"ORD-{oid}", f"{y}-{m:02d}-{day:02d}", name, cat, qty, price, qty * price])
OUT.parent.mkdir(parents=True, exist_ok=True)
with open(OUT, "w", newline="") as f:
    w = csv.writer(f); w.writerow(["order_id", "date", "product", "category", "qty", "unit_price_pkr", "amount_pkr"]); w.writerows(rows)
print(f"wrote {len(rows)} rows to {OUT}")
