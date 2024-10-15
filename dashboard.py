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

    # Load data for Bobby Aquatic 1
    folder_path_1 = "./data/Bobby Aquatic 1"
    sheet_name_1 = 'Penjualan'
    penjualan_data_1 = load_data_1(folder_path_1, sheet_name_1)
    daily_profit_1, hw_forecast_future_1 = forecast_profit_1(penjualan_data_1)

    # Load data for Bobby Aquatic 2
    folder_path_2 = "./data/Bobby Aquatic 2"
    sheet_name_2 = 'Penjualan'
    penjualan_data_2 = load_data_2(folder_path_2, sheet_name_2)
    daily_profit_2, hw_forecast_future_2 = forecast_profit_2(penjualan_data_2)

    # Ensure TANGGAL columns are correctly named and formatted
    # Assuming 'daily_profit_1' and 'daily_profit_2' have a TANGGAL column named 'TANGGAL' or are indexed by TANGGAL
    if 'TANGGAL' in daily_profit_1.columns:
        daily_profit_1['TANGGAL'] = pd.to_datetime(daily_profit_1['TANGGAL'])
        daily_profit_2['TANGGAL'] = pd.to_datetime(daily_profit_2['TANGGAL'])
    else:
        # If the data uses an index as the TANGGAL
        daily_profit_1 = daily_profit_1.reset_index().rename(columns={'index': 'TANGGAL'})
        daily_profit_2 = daily_profit_2.reset_index().rename(columns={'index': 'TANGGAL'})

    # Combine data
    combined_profit = pd.concat([daily_profit_1, daily_profit_2]).groupby('TANGGAL').sum().reset_index()
    combined_forecast = hw_forecast_future_1.add(hw_forecast_future_2, fill_value=0)

    # Filter for selecting branches
    branch_option = st.selectbox(
        "Pilih Cabang", 
        ["Gabungan", "Bobby Aquatic 1", "Bobby Aquatic 2"],
        index=0
    )

    # Generic function to display dashboards
    def show_combined_dashboard(profit_data, forecast_data, title):
        st.header(title)
        # Visualization code goes here (e.g., line charts, bar charts)
        st.line_chart(profit_data.set_index('TANGGAL')['Profit'], width=0, height=300)  # Example line chart
        st.write(forecast_data)  # Example table display

    # Display the selected data
    if branch_option == "Gabungan":
        show_combined_dashboard(combined_profit, combined_forecast, "Penjualan Gabungan Bobby Aquatic 1 & 2")
    elif branch_option == "Bobby Aquatic 1":
        show_combined_dashboard(daily_profit_1, hw_forecast_future_1, "Penjualan Bobby Aquatic 1")
    else:
        show_combined_dashboard(daily_profit_2, hw_forecast_future_2, "Penjualan Bobby Aquatic 2")

elif st.session_state.page == "product":
    st.header("🔍 Segmentasi Produk Bobby Aquatic")

    # Tabs for Bobby Aquatic 1 and 2
    tab1, tab2 = st.tabs(["Bobby Aquatic 1", "Bobby Aquatic 2"])

    # Bobby Aquatic 1 clustering dashboard
    with tab1:
        st.header("Segmentasi Produk Bobby Aquatic 1")
        folder_path_1 = "./data/Bobby Aquatic 1"
        sheet_name_1 = 'Penjualan'
        cluster_data_1 = load_cluster_data_1(folder_path_1, sheet_name_1)
        show_cluster_dashboard_1(cluster_data_1, key_suffix='cabang1')

    # Bobby Aquatic 2 clustering dashboard
    with tab2:
        st.header("Segmentasi Produk Bobby Aquatic 2")
        folder_path_2 = "./data/Bobby Aquatic 2"
        sheet_name_2 = 'Penjualan'
        cluster_data_2 = load_cluster_data_2(folder_path_2, sheet_name_2)
        show_cluster_dashboard_2(cluster_data_2, key_suffix='cabang2')

# Footer section
st.markdown("<div class='footer'>© 2024 Bobby Aquatic. All rights reserved.</div>", unsafe_allow_html=True)
