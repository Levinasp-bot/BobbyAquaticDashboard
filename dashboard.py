import streamlit as st
import pandas as pd
from sales_forecast1 import load_all_excel_files as load_data_1, forecast_profit as forecast_profit_1, show_dashboard
from sales_forecast2 import load_all_excel_files as load_data_2, forecast_profit as forecast_profit_2
from product_clustering import load_all_excel_files as load_cluster_data_1, show_dashboard as show_cluster_dashboard_1
from product_clustering2 import load_all_excel_files as load_cluster_data_2, show_dashboard as show_cluster_dashboard_2

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

    # Filter cabang di bawah header
    branch_selection = st.multiselect(
        "Pilih cabang untuk ditampilkan:",
        options=["Bobby Aquatic 1", "Bobby Aquatic 2"],
        default=["Bobby Aquatic 1", "Bobby Aquatic 2"]
    )

    daily_profit_combined = None
    hw_forecast_future_combined = None

    # Load sales data based on selected branches
    if "Bobby Aquatic 1" in branch_selection:
        folder_path_1 = "./data/Bobby Aquatic 1"
        sheet_name_1 = 'Penjualan'
        penjualan_data_1 = load_data_1(folder_path_1, sheet_name_1)

        # Forecast profit based on the selected branch data
        daily_profit_1, hw_forecast_future_1 = forecast_profit_1(penjualan_data_1)

    if "Bobby Aquatic 2" in branch_selection:
        folder_path_2 = "./data/Bobby Aquatic 2"
        sheet_name_2 = 'Penjualan'
        penjualan_data_2 = load_data_2(folder_path_2, sheet_name_2)

        # Forecast profit based on the selected branch data
        daily_profit_2, hw_forecast_future_2 = forecast_profit_2(penjualan_data_2)

    # Combine profits if both branches are selected
    if "Bobby Aquatic 1" in branch_selection and "Bobby Aquatic 2" in branch_selection:
        combined_penjualan_data = pd.concat([penjualan_data_1, penjualan_data_2], ignore_index=True)

        # Forecast profit based on combined data
        daily_profit_combined, hw_forecast_future_combined = forecast_profit_1(combined_penjualan_data)

# Show dashboard based on selections
    if "Bobby Aquatic 1" in branch_selection and "Bobby Aquatic 2" in branch_selection:
    # Show dashboard for the combined data
        show_dashboard(daily_profit_1, hw_forecast_future_1, daily_profit_2, hw_forecast_future_2, key_suffix='combined')
    elif "Bobby Aquatic 1" in branch_selection:
    # Show dashboard for Bobby Aquatic 1
        show_dashboard(daily_profit_1, hw_forecast_future_1, None, None, key_suffix='cabang1')  # Pass None for second branch
    elif "Bobby Aquatic 2" in branch_selection:
    # Show dashboard for Bobby Aquatic 2
        show_dashboard(None, None, daily_profit_2, hw_forecast_future_2, key_suffix='cabang2')  # Pass None for first branch


elif st.session_state.page == "product":
    st.header("🔍 Segmentasi Produk Bobby Aquatic")

    # Tabs for Bobby Aquatic 1 and 2
    tab1, tab2 = st.tabs(["Bobby Aquatic 1", "Bobby Aquatic 2"])

    # Bobby Aquatic 1 clustering dashboard
    with tab1:
        st.header("Segmentasi Produk Bobby Aquatic 1")

        # Load data for Bobby Aquatic 1
        folder_path_1 = "./data/Bobby Aquatic 1"
        sheet_name_1 = 'Penjualan'
        cluster_data_1 = load_cluster_data_1(folder_path_1, sheet_name_1)

        # Show cluster dashboard
        show_cluster_dashboard_1(cluster_data_1, key_suffix='cabang1')

    # Bobby Aquatic 2 clustering dashboard
    with tab2:
        st.header("Segmentasi Produk Bobby Aquatic 2")

        # Load data for Bobby Aquatic 2
        folder_path_2 = "./data/Bobby Aquatic 2"
        sheet_name_2 = 'Penjualan'
        cluster_data_2 = load_cluster_data_2(folder_path_2, sheet_name_2)

        # Show cluster dashboard
        show_cluster_dashboard_2(cluster_data_2, key_suffix='cabang2')

# Footer section
st.markdown("<div class='footer'>© 2024 Bobby Aquatic. All rights reserved.</div>", unsafe_allow_html=True)
