
import pandas as pd
import numpy as np
from datetime import datetime, timedelta

def generate_report(products_file, orders_file, opening_stock_file=None):

    products = pd.read_excel(products_file)
    orders = pd.read_excel(orders_file)

    # Inventory
    inventory = (
        products.groupby(["ID","Handle","Title"], dropna=False)["Variant Inventory Qty"]
        .sum()
        .reset_index()
    )

    inventory.columns = ["Product ID","Handle","Product Name","Current Stock"]

    # Orders
    orders = orders[orders["Line: Product ID"].notna()].copy()

    if "Payment: Status" in orders.columns:
        orders = orders[orders["Payment: Status"].astype(str).str.lower() != "voided"]

    orders["Created At"] = pd.to_datetime(orders["Created At"], errors="coerce")

    last30 = datetime.now() - timedelta(days=30)
    last90 = datetime.now() - timedelta(days=90)

    orders30 = orders[orders["Created At"] >= last30]
    orders90 = orders[orders["Created At"] >= last90]

    sold30 = (
        orders30.groupby("Line: Product ID")["Line: Quantity"]
        .sum()
        .reset_index()
    )
    sold30.columns = ["Product ID","Sold Last 30 Days"]

    report = inventory.merge(sold30, on="Product ID", how="left")
    report["Sold Last 30 Days"] = report["Sold Last 30 Days"].fillna(0)

    report["Daily Average Sales"] = report["Sold Last 30 Days"] / 30

    report["Cover Days"] = np.where(
        report["Daily Average Sales"] > 0,
        report["Current Stock"] / report["Daily Average Sales"],
        9999
    )

    # Peak week
    orders90["Week"] = orders90["Created At"].dt.to_period("W").astype(str)

    peak = (
        orders90.groupby(["Line: Product ID","Week"])["Line: Quantity"]
        .sum()
        .reset_index()
        .groupby("Line: Product ID")["Line: Quantity"]
        .max()
        .reset_index()
    )

    peak.columns = ["Product ID","Peak Week Sales"]

    report = report.merge(peak, on="Product ID", how="left")
    report["Peak Week Sales"] = report["Peak Week Sales"].fillna(0)

    report["Cover Weeks"] = np.where(
        report["Peak Week Sales"] > 0,
        report["Current Stock"] / report["Peak Week Sales"],
        9999
    )

    report["Stock Category"] = np.select(
        [
            report["Current Stock"] > 20,
            report["Current Stock"].between(10,20)
        ],
        [
            "Above 20",
            "Stock 10-20"
        ],
        default="Below 10"
    )

    full_inventory = report.sort_values("Current Stock", ascending=False)
    above20 = report[report["Current Stock"] > 20]
    stock1020 = report[report["Current Stock"].between(10,20)]
    below10 = report[report["Current Stock"] < 10]

    fast_moving = report[report["Sold Last 30 Days"] > 0].sort_values(
        "Sold Last 30 Days", ascending=False
    )

    peak_week_risk = report[report["Cover Weeks"] < 4].sort_values(
        "Cover Weeks"
    )

    # Broken models
    broken_rows = []

    for (pid, handle, title), grp in products.groupby(["ID","Handle","Title"]):

        low = grp[grp["Variant Inventory Qty"] <= 3]

        if len(low) > 0:

            variants = []

            for _, r in low.iterrows():

                vals = [
                    str(x)
                    for x in [
                        r.get("Option1 Value"),
                        r.get("Option2 Value"),
                        r.get("Option3 Value")
                    ]
                    if pd.notna(x)
                ]

                variants.append(" / ".join(vals))

            broken_rows.append({
                "Product ID": pid,
                "Handle": handle,
                "Product Name": title,
                "Low Stock Variants": " | ".join(variants),
                "Low Variant Count": len(low)
            })

    broken_models = pd.DataFrame(broken_rows)

    # Sell Through
    sell_through = pd.DataFrame()

    if opening_stock_file is not None:

        raw = pd.read_excel(opening_stock_file, header=None)

        opening = raw.iloc[2:].copy()
        opening.columns = [
            "#",
            "Product Name",
            "Handle",
            "Current Stock",
            "Sold Qty",
            "Opening Stock"
        ]

        opening = opening.dropna(subset=["Handle"])

        opening["Opening Stock"] = pd.to_numeric(opening["Opening Stock"], errors="coerce")
        opening["Current Stock"] = pd.to_numeric(opening["Current Stock"], errors="coerce")
        opening["Sold Qty"] = pd.to_numeric(opening["Sold Qty"], errors="coerce")

        opening["Sell Through %"] = (
            opening["Sold Qty"] / opening["Opening Stock"] * 100
        )

        opening["Remaining %"] = (
            opening["Current Stock"] / opening["Opening Stock"] * 100
        )

        sell_through = opening

    output = "Inventory_Report.xlsx"

    with pd.ExcelWriter(output, engine="openpyxl") as writer:

        full_inventory.to_excel(writer, sheet_name="Full_Inventory_Report", index=False)
        above20.to_excel(writer, sheet_name="Above_20", index=False)
        stock1020.to_excel(writer, sheet_name="Stock_10_20", index=False)
        below10.to_excel(writer, sheet_name="Below_10", index=False)
        fast_moving.to_excel(writer, sheet_name="Fast_Moving", index=False)
        peak_week_risk.to_excel(writer, sheet_name="Peak_Week_Risk", index=False)
        broken_models.to_excel(writer, sheet_name="Broken_Models", index=False)

        if len(sell_through) > 0:
            sell_through.to_excel(writer, sheet_name="Sell_Through_Analysis", index=False)

    return output
