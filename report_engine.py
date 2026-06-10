import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime, timedelta

def generate_report(products_file, orders_file, opening_stock_file=None):

```
products = pd.read_excel(products_file)
orders = pd.read_excel(orders_file)

# =========================
# PRODUCTS
# =========================

inventory = (
    products.groupby(
        ["ID", "Handle", "Title"],
        dropna=False
    )["Variant Inventory Qty"]
    .sum()
    .reset_index()
)

inventory.rename(
    columns={
        "ID": "Product ID",
        "Handle": "Handle",
        "Title": "Product Name",
        "Variant Inventory Qty": "Current Stock"
    },
    inplace=True
)

# =========================
# ORDERS
# =========================

orders = orders[
    orders["Line: Product ID"].notna()
]

if "Payment: Status" in orders.columns:
    orders = orders[
        orders["Payment: Status"].astype(str).str.lower()
        != "voided"
    ]

orders["Created At"] = pd.to_datetime(
    orders["Created At"],
    errors="coerce"
)

last_30_days = datetime.now() - timedelta(days=30)

orders_30 = orders[
    orders["Created At"] >= last_30_days
]

sold30 = (
    orders_30.groupby("Line: Product ID")["Line: Quantity"]
    .sum()
    .reset_index()
)

sold30.columns = [
    "Product ID",
    "Sold Last 30 Days"
]

report = inventory.merge(
    sold30,
    on="Product ID",
    how="left"
)

report["Sold Last 30 Days"] = (
    report["Sold Last 30 Days"]
    .fillna(0)
)

report["Daily Average Sales"] = (
    report["Sold Last 30 Days"] / 30
)

report["Cover Days"] = np.where(
    report["Daily Average Sales"] > 0,
    report["Current Stock"]
    / report["Daily Average Sales"],
    9999
)
```
