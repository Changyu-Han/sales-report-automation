"""Generate deliberately messy sample sales data from three fictional stores.

Each store exports data differently (different column names, date formats,
currency formatting, languages), which mirrors what real small businesses send.
Run:  python scripts/generate_sample_data.py
"""
from pathlib import Path
import random

import pandas as pd

RAW_DIR = Path(__file__).resolve().parent.parent / "data" / "raw"
PRODUCTS = {
    "Brown Sugar Milk Tea": 7.5,
    "Taro Milk Tea": 7.0,
    "Mango Smoothie": 8.0,
    "Matcha Latte": 7.8,
    "Lemon Green Tea": 6.0,
    "Cheese Foam Oolong": 8.5,
}
DAYS = pd.date_range("2026-07-01", "2026-09-30", freq="D")


def _rows(n: int, seed: int):
    rng = random.Random(seed)
    for _ in range(n):
        name = rng.choice(list(PRODUCTS))
        yield rng.choice(DAYS), name, rng.randint(1, 6), PRODUCTS[name]


def order_ids(prefix: str, n: int):
    return [f"{prefix}{i:05d}" for i in range(1, n + 1)]


def messy_name(name: str, rng: random.Random) -> str:
    """Random casing / stray whitespace, like hand-typed POS entries."""
    variant = rng.choice([name, name.lower(), name.upper(), f"  {name} "])
    return variant


def store_a(rng):  # CSV, ISO dates, "$" prices
    rows = [
        {"Order #": oid, "Date": d.strftime("%Y-%m-%d"), "Product": messy_name(p, rng),
         "Qty": q, "Unit Price": f"${price:.2f}"}
        for oid, (d, p, q, price) in zip(order_ids("BD", 400), _rows(400, 1))
    ]
    df = pd.DataFrame(rows)
    df = pd.concat([df, df.sample(15, random_state=1)])  # duplicate exports
    df.loc[df.sample(5, random_state=2).index, "Qty"] = None  # missing values
    df.to_csv(RAW_DIR / "store_a_bondi.csv", index=False)


def store_b(rng):  # Excel, DD/MM/YYYY dates, snake_case headers
    rows = [
        {"receipt_no": oid, "order_date": d.strftime("%d/%m/%Y"), "item": messy_name(p, rng),
         "quantity": q, "price": price}
        for oid, (d, p, q, price) in zip(order_ids("CW", 350), _rows(350, 2))
    ]
    rows.append({"receipt_no": None, "order_date": "", "item": "", "quantity": None, "price": None})  # blank line
    rows.append({"receipt_no": "CW00351", "order_date": "15/08/2026", "item": "Taro Milk Tea", "quantity": -2, "price": 7.0})  # refund
    pd.DataFrame(rows).to_excel(RAW_DIR / "store_b_chatswood.xlsx", index=False)


def store_c(rng):  # CSV with POS-style headers, slash dates, comma-formatted numbers
    rows = [
        {"Order No": oid, "Transaction Date": d.strftime("%Y/%m/%d"), "SKU Name": messy_name(p, rng),
         "Units": q, "Unit_Price": f"{price:,.2f}"}
        for oid, (d, p, q, price) in zip(order_ids("BW", 300), _rows(300, 3))
    ]
    pd.DataFrame(rows).to_csv(RAW_DIR / "store_c_burwood.csv", index=False, encoding="utf-8-sig")


if __name__ == "__main__":
    RAW_DIR.mkdir(parents=True, exist_ok=True)
    rng = random.Random(42)
    store_a(rng)
    store_b(rng)
    store_c(rng)
    print(f"Sample data written to {RAW_DIR}")
