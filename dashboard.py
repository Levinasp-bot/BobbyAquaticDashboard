import streamlit as st
import pandas as pd
from sales_forecast1 import load_all_excel_files as load_data_1, forecast_profit as forecast_profit_1, show_dashboard
from sales_forecast2 import load_all_excel_files as load_data_2, forecast_profit as forecast_profit_2
from product_clustering import load_all_excel_files as load_cluster_data_1, show_dashboard as show_cluster_dashboard_1
from product_clustering2 import load_all_excel_files as load_cluster_data_2, show_dashboard as show_cluster_dashboard_2

st.set_page_config(page_title="Bobby Aquatic Dashboard", layout="wide")

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

if 'page' not in st.session_state:
    st.session_state.page = 'sales'

def switch_page(page_name):
    st.session_state.page = page_name

with st.sidebar:
    st.title("Dashboard Navigation")
    st.markdown("### Pilih Halaman:")
    if st.button('📊 Penjualan', key="sales_button"):
        switch_page("sales")
    if st.button('📦 Produk', key="product_button"):
        switch_page("product")

if st.session_state.page == "sales":
    st.header("📈 Dashboard Penjualan Bobby Aquatic")

    branch_selection = st.multiselect(
        "Pilih cabang untuk ditampilkan:",
        options=["Bobby Aquatic 1", "Bobby Aquatic 2"],
        default=["Bobby Aquatic 1", "Bobby Aquatic 2"]
    )

    if "Bobby Aquatic 1" in branch_selection:
        folder_path_1 = "./data/Bobby Aquatic 1"
        sheet_name_1 = 'Penjualan'
        penjualan_data_1 = load_data_1(folder_path_1, sheet_name_1)

        daily_profit_1, fitted_values_1, test_1, test_forecast_1, hw_forecast_future_1 = forecast_profit_1(penjualan_data_1)

    if "Bobby Aquatic 2" in branch_selection:
        folder_path_2 = "./data/Bobby Aquatic 2"
        sheet_name_2 = 'Penjualan'
        penjualan_data_2 = load_data_2(folder_path_2, sheet_name_2)

        daily_profit_2, fitted_values_2, test_2, test_forecast_2, hw_forecast_future_2 = forecast_profit_2(penjualan_data_2)

    if "Bobby Aquatic 1" in branch_selection and "Bobby Aquatic 2" in branch_selection:
        combined_penjualan_data = pd.concat([penjualan_data_1, penjualan_data_2], ignore_index=True)

        daily_profit_combined, fitted_values_combined, test_combined, test_forecast_combined, hw_forecast_future_combined = forecast_profit_1(combined_penjualan_data)

    if "Bobby Aquatic 1" in branch_selection and "Bobby Aquatic 2" in branch_selection:
        show_dashboard(
            daily_profit_1, fitted_values_1, test_1, test_forecast_1, hw_forecast_future_1, 
            daily_profit_2, fitted_values_2, test_2, test_forecast_2, hw_forecast_future_2,
            key_suffix='combined'
        )
    elif "Bobby Aquatic 1" in branch_selection:
        show_dashboard(
            daily_profit_1, fitted_values_1, test_1, test_forecast_1, hw_forecast_future_1, 
            None, None, None, None, None, key_suffix='cabang1'
        )
    elif "Bobby Aquatic 2" in branch_selection:
        show_dashboard(
            None, None, None, None, None,
            daily_profit_2, fitted_values_2, test_2, test_forecast_2, hw_forecast_future_2, 
            key_suffix='cabang2'
        )

elif st.session_state.page == "product":
    st.header("🔍 Segmentasi Produk Bobby Aquatic")

    tab1, tab2 = st.tabs(["Bobby Aquatic 1", "Bobby Aquatic 2"])

    with tab1:
        st.header("Segmentasi Produk Bobby Aquatic 1")

        folder_path_1 = "./data/Bobby Aquatic 1"
        sheet_name_1 = 'Penjualan'
        cluster_data_1 = load_cluster_data_1(folder_path_1, sheet_name_1)

        show_cluster_dashboard_1(cluster_data_1, key_suffix='cabang1')

    with tab2:
        st.header("Segmentasi Produk Bobby Aquatic 2")

        folder_path_2 = "./data/Bobby Aquatic 2"
        sheet_name_2 = 'Penjualan'
        cluster_data_2 = load_cluster_data_2(folder_path_2, sheet_name_2)

        show_cluster_dashboard_2(cluster_data_2, key_suffix='cabang2')

st.markdown("<div class='footer'>© 2024 Bobby Aquatic. All rights reserved.</div>", unsafe_allow_html=True)
