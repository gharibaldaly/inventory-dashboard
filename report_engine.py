import pandas as pd
import numpy as np
from datetime import datetime, timedelta


def generate_report(products_file, orders_file, opening_stock_file=None):

    products = pd.read_excel(products_file)
    orders = pd.read_excel(orders_file)

    # =========================
    # INVENTORY
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
    ].copy()

    if "Payment: Status" in orders.columns:

        orders = orders[
            orders["Payment: Status"]
            .astype(str)
            .str.lower()
            != "voided"
        ]

    orders["Created At"] = pd.to_datetime(
        orders["Created At"],
        errors="coerce"
    )

    # =========================
    # SOLD LAST 30 DAYS
    # =========================

    last_30 = datetime.now() - timedelta(days=30)

    orders30 = orders[
        orders["Created At"] >= last_30
    ]

    sold30 = (
        orders30.groupby("Line: Product ID")["Line: Quantity"]
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

    # =========================
    # DAILY AVG SALES
    # =========================

    report["Daily Average Sales"] = (
        report["Sold Last 30 Days"] / 30
    )

    report["Cover Days"] = np.where(
        report["Daily Average Sales"] > 0,
        report["Current Stock"]
        / report["Daily Average Sales"],
        9999
    )

    # =========================
    # PEAK WEEK SALES
    # =========================

    last_90 = datetime.now() - timedelta(days=90)

    orders90 = orders[
        orders["Created At"] >= last_90
    ].copy()

    orders90["Week"] = (
        orders90["Created At"]
        .dt.to_period("W")
        .astype(str)
    )

    weekly_sales = (
        orders90.groupby(
            ["Line: Product ID", "Week"]
        )["Line: Quantity"]
        .sum()
        .reset_index()
    )

    peak_week = (
        weekly_sales.groupby(
            "Line: Product ID"
        )["Line: Quantity"]
        .max()
        .reset_index()
    )

    peak_week.columns = [
        "Product ID",
        "Peak Week Sales"
    ]

    report = report.merge(
        peak_week,
        on="Product ID",
        how="left"
    )

    report["Peak Week Sales"] = (
        report["Peak Week Sales"]
        .fillna(0)
    )

    report["Cover Weeks"] = np.where(
        report["Peak Week Sales"] > 0,
        report["Current Stock"]
        / report["Peak Week Sales"],
        9999
    )

    # =========================
    # STOCK CATEGORY
    # =========================

    report["Stock Category"] = np.select(
        [
            report["Current Stock"] > 20,
            report["Current Stock"].between(10, 20)
        ],
        [
            "Above 20",
            "Stock 10-20"
        ],
        default="Below 10"
    )

    # =========================
    # REPORTS
    # =========================

    full_inventory = report.sort_values(
        "Current Stock",
        ascending=False
    )

    above20 = report[
        report["Current Stock"] > 20
    ]

    stock10_20 = report[
        report["Current Stock"].between(10, 20)
    ]

    below10 = report[
        report["Current Stock"] < 10
    ]

    fast_moving = report[
        report["Sold Last 30 Days"] > 0
    ].sort_values(
        "Sold Last 30 Days",
        ascending=False
    )

    peak_week_risk = report[
        report["Cover Weeks"] < 4
    ].sort_values(
        "Cover Weeks",
        ascending=True
    )

    # =========================
    # EXPORT
    # =========================

    output_file = "Inventory_Report.xlsx"

    with pd.ExcelWriter(
        output_file,
        engine="openpyxl"
    ) as writer:

        full_inventory.to_excel(
            writer,
            sheet_name="Full_Inventory_Report",
            index=False
        )

        above20.to_excel(
            writer,
            sheet_name="Above_20",
            index=False
        )

        stock10_20.to_excel(
            writer,
            sheet_name="Stock_10_20",
            index=False
        )

        below10.to_excel(
            writer,
            sheet_name="Below_10",
            index=False
        )

        fast_moving.to_excel(
            writer,
            sheet_name="Fast_Moving",
            index=False
        )

        peak_week_risk.to_excel(
            writer,
            sheet_name="Peak_Week_Risk",
            index=False
        )

    return output_file
