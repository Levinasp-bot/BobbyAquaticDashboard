import streamlit as st
from sales_forecast1 import load_all_excel_files as load_data_1, forecast_profit as forecast_profit_1
from sales_forecast2 import load_all_excel_files as load_data_2, forecast_profit as forecast_profit_2
from product_clustering import load_all_excel_files as load_cluster_data_1, show_dashboard as show_cluster_dashboard_1
from product_clustering2 import load_all_excel_files as load_cluster_data_2, show_dashboard as show_cluster_dashboard_2
import pandas as pd

# Set page configuration
st.set_page_config(page_title="Bobby Aquatic Dashboard", layout="wide")

# Custom CSS for styling
st.markdown("""
<style>
    .sidebar .sidebar-content {
        background-color: #f0f2f5;
        padding: 20px;
        border-right: 2px solid #2A7B5C;
    }
    .header {
        font-size: 1.5em;
        color: #2A7B5C;
    }
    .footer {
        font-size: 0.8em;
        color: #7D7D7D;
        text-align: center;
        margin-top: 20px;
    }
    .stTabs .stTabs-header {
        justify-content: center;
    }
</style>
""", unsafe_allow_html=True)

# Initialize session state for page navigation
if 'page' not in st.session_state:
    st.session_state.page = 'sales'

def switch_page(page_name):
    st.session_state.page = page_name

# Sidebar for navigation
with st.sidebar:
    st.title("Dashboard Navigation")
    st.markdown("### Pilih Halaman:")
    if st.button('📊 Penjualan', key="sales_button"):
        switch_page("sales")
    if st.button('📦 Produk', key="product_button"):
        switch_page("product")

# Display content based on selected page
if st.session_state.page == "sales":
    st.header("📈 Dashboard Penjualan Bobby Aquatic")

    # Load data for Bobby Aquatic 1 and 2
    folder_path_1 = "./data/Bobby Aquatic 1"
    sheet_name_1 = 'Penjualan'
    penjualan_data_1 = load_data_1(folder_path_1, sheet_name_1)

    folder_path_2 = "./data/Bobby Aquatic 2"
    sheet_name_2 = 'Penjualan'
    penjualan_data_2 = load_data_2(folder_path_2, sheet_name_2)

    # Combine sales data
    combined_penjualan_data = pd.concat([penjualan_data_1, penjualan_data_2]).groupby('Product').sum().reset_index()

    # Forecast combined data
    combined_daily_profit, combined_hw_forecast_future = forecast_profit_1(combined_penjualan_data)  # Use the same forecast function

    # Show combined sales dashboard
    st.header("Dashboard Penjualan Gabungan Bobby Aquatic")
    show_dashboard_1(combined_daily_profit, combined_hw_forecast_future, key_suffix='gabungan')

elif st.session_state.page == "product":
    st.header("🔍 Segmentasi Produk Bobby Aquatic")

    # Load and combine clustering data for Bobby Aquatic 1 and 2
    cluster_data_1 = load_cluster_data_1(folder_path_1, sheet_name_1)
    cluster_data_2 = load_cluster_data_2(folder_path_2, sheet_name_2)

    combined_cluster_data = pd.concat([cluster_data_1, cluster_data_2]).groupby('Product').sum().reset_index()

    # Show combined clustering dashboard
    show_cluster_dashboard_1(combined_cluster_data, key_suffix='gabungan')

# Footer section
st.markdown("<div class='footer'>© 2024 Bobby Aquatic. All rights reserved.</div>", unsafe_allow_html=True)
