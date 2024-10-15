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
    folder_path_2 = "./data/Bobby Aquatic 2"
    sheet_name = 'Penjualan'
    penjualan_data_1 = load_data_1(folder_path_1, sheet_name)
    penjualan_data_2 = load_data_2(folder_path_2, sheet_name)

    # Combine sales data
    combined_data = pd.concat([penjualan_data_1, penjualan_data_2], ignore_index=True)

    # Forecast data for combined branches
    daily_profit_1, hw_forecast_future_1 = forecast_profit_1(penjualan_data_1)
    daily_profit_2, hw_forecast_future_2 = forecast_profit_2(penjualan_data_2)

    # Option to view combined data or separate
    st.markdown("### Pilih Tipe Data:")
    view_type = st.radio("Tampilkan data sebagai:", ("Gabungan", "Terpisah"))

    if view_type == "Gabungan":
        # Sum the daily profits for a combined view
        combined_daily_profit = daily_profit_1.add(daily_profit_2, fill_value=0)
        combined_forecast = hw_forecast_future_1.add(hw_forecast_future_2, fill_value=0)

        # Show combined dashboard
        st.header("Dashboard Penjualan Gabungan Bobby Aquatic")
        st.line_chart(combined_daily_profit, title="Profit Harian Gabungan")
        st.line_chart(combined_forecast, title="Peramalan Profit Gabungan")
    else:
        # Show separate dashboards in a single tab
        st.subheader("Dashboard Penjualan Bobby Aquatic 1")
        st.line_chart(daily_profit_1, title="Profit Harian Bobby Aquatic 1")
        st.line_chart(hw_forecast_future_1, title="Peramalan Profit Bobby Aquatic 1")

        st.subheader("Dashboard Penjualan Bobby Aquatic 2")
        st.line_chart(daily_profit_2, title="Profit Harian Bobby Aquatic 2")
        st.line_chart(hw_forecast_future_2, title="Peramalan Profit Bobby Aquatic 2")

elif st.session_state.page == "product":
    st.header("🔍 Segmentasi Produk Bobby Aquatic")

    # Tabs for Bobby Aquatic 1 and 2
    tab1, tab2 = st.tabs(["Bobby Aquatic 1", "Bobby Aquatic 2"])

    # Bobby Aquatic 1 clustering dashboard
    with tab1:
        st.header("Segmentasi Produk Bobby Aquatic 1")

        # Load data for Bobby Aquatic 1
        folder_path_1 = "./data/Bobby Aquatic 1"
        cluster_data_1 = load_cluster_data_1(folder_path_1, sheet_name)

        # Show cluster dashboard
        show_cluster_dashboard_1(cluster_data_1, key_suffix='cabang1')

    # Bobby Aquatic 2 clustering dashboard
    with tab2:
        st.header("Segmentasi Produk Bobby Aquatic 2")

        # Load data for Bobby Aquatic 2
        folder_path_2 = "./data/Bobby Aquatic 2"
        cluster_data_2 = load_cluster_data_2(folder_path_2, sheet_name)

        # Show cluster dashboard
        show_cluster_dashboard_2(cluster_data_2, key_suffix='cabang2')

# Footer section
st.markdown("<div class='footer'>© 2024 Bobby Aquatic. All rights reserved.</div>", unsafe_allow_html=True)
