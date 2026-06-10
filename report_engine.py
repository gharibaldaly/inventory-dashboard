import pandas as pd

def generate_report(products_file, orders_file, opening_stock_file=None):

    products = pd.read_excel(products_file)

    output = "Inventory_Report.xlsx"

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        products.head(50).to_excel(
            writer,
            sheet_name="Test",
            index=False
        )

    return output
