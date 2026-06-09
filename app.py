import streamlit as st

st.set_page_config(
    page_title="Inventory Dashboard",
    layout="wide"
)

st.title("📦 Inventory Analytics Dashboard")

products_file = st.file_uploader(
    "Upload Products.xlsx",
    type=["xlsx"]
)

orders_file = st.file_uploader(
    "Upload Orders.xlsx",
    type=["xlsx"]
)

opening_stock_file = st.file_uploader(
    "Upload Opening Stock.xlsx (Optional)",
    type=["xlsx"]
)

if st.button("Generate Report"):

    if products_file is None:
        st.error("Please upload Products file")
        st.stop()

    if orders_file is None:
        st.error("Please upload Orders file")
        st.stop()

    st.success("Files uploaded successfully ✅")

    st.info(
        "Report Engine will be connected in the next step."
    )
