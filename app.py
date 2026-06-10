import streamlit as st
from report_engine import generate_report

st.set_page_config(
    page_title="Inventory Analytics Dashboard",
    page_icon="📦",
    layout="wide"
)

st.title("📦 Inventory Analytics Dashboard")
st.markdown("Upload Matrixify files and generate the inventory report.")

col1, col2, col3 = st.columns(3)

with col1:
    products_file = st.file_uploader(
        "Upload Products.xlsx",
        type=["xlsx"],
        key="products"
    )

with col2:
    orders_file = st.file_uploader(
        "Upload Orders.xlsx",
        type=["xlsx"],
        key="orders"
    )

with col3:
    opening_stock_file = st.file_uploader(
        "Upload Opening Stock.xlsx (Optional)",
        type=["xlsx"],
        key="opening_stock"
    )

st.divider()

if st.button("🚀 Generate Report", use_container_width=True):

    if products_file is None:
        st.error("Please upload Products.xlsx")
        st.stop()

    if orders_file is None:
        st.error("Please upload Orders.xlsx")
        st.stop()

    try:

        with st.spinner("Generating Inventory Report..."):

            report_file = generate_report(
                products_file,
                orders_file,
                opening_stock_file
            )

        st.success("Report generated successfully ✅")

        with open(report_file, "rb") as file:

            st.download_button(
                label="📥 Download Inventory_Report.xlsx",
                data=file,
                file_name="Inventory_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

    except Exception as e:

        st.error(f"Error: {str(e)}")
