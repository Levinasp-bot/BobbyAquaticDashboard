import streamlit as st
from sales_forecast1 import load_all_excel_files as load_data_1, forecast_profit as forecast_profit_1
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
        border-right: 2px solid #2A7B5C;  /* Added border for better visibility */
    }
    .header {
        font-size: 1.5em;
        color: #2A7B5C;
    }
    .footer {
        font-size: 0.8em;
        color: #7D7D7D;
        text-align: center;  /* Center the copyright text */
        margin-top: 20px;  /* Added margin for spacing */
    }
    .stTabs .stTabs-header {
        justify-content: center;  /* Center tab headers */
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

    # Gabungkan data penjualan
    penjualan_data_combined = penjualan_data_1 + penjualan_data_2

    # Forecast data gabungan
    daily_profit_combined, hw_forecast_future_combined = forecast_profit_1(penjualan_data_combined)

    # Show combined dashboard
    show_dashboard_1(daily_profit_combined, hw_forecast_future_combined, key_suffix='gabungan')

elif st.session_state.page == "product":
    st.header("🔍 Segmentasi Produk Bobby Aquatic")

    # Load data for Bobby Aquatic 1 and 2
    folder_path_1 = "./data/Bobby Aquatic 1"
    sheet_name_1 = 'Penjualan'
    cluster_data_1 = load_cluster_data_1(folder_path_1, sheet_name_1)

    folder_path_2 = "./data/Bobby Aquatic 2"
    sheet_name_2 = 'Penjualan'
    cluster_data_2 = load_cluster_data_2(folder_path_2, sheet_name_2)

    # Gabungkan data cluster jika diperlukan

    # Show cluster dashboard (could be adjusted for combined data if needed)
    show_cluster_dashboard_1(cluster_data_1, key_suffix='cabang1')
    show_cluster_dashboard_2(cluster_data_2, key_suffix='cabang2')

# Footer section
st.markdown("<div class='footer'>© 2024 Bobby Aquatic. All rights reserved.</div>", unsafe_allow_html=True)
