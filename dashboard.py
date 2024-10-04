import streamlit as st
from sales_forecast1 import load_all_excel_files as load_data_1, forecast_profit as forecast_profit_1, show_dashboard as show_dashboard_1
from sales_forecast2 import load_all_excel_files as load_data_2, forecast_profit as forecast_profit_2, show_dashboard as show_dashboard_2
from product_clustering import load_all_excel_files as load_cluster_data_1, show_dashboard as show_cluster_dashboard_1
from product_clustering2 import load_all_excel_files as load_cluster_data_2, show_dashboard as show_cluster_dashboard_2

# Mengatur layout agar wide mode
st.set_page_config(layout="wide")

# Inisialisasi session state untuk halaman jika belum ada
if 'page' not in st.session_state:
    st.session_state.page = 'sales'  # Set default page

# Fungsi untuk mengganti halaman
def switch_page(page_name):
    st.session_state.page = page_name

# Sidebar untuk navigasi antara Sales dan Product
with st.sidebar:
    if st.button('Penjualan', key="sales_button"):
        switch_page("sales")
    if st.button('Produk', key="product_button"):
        switch_page("product")

# Menampilkan konten berdasarkan halaman yang dipilih
if st.session_state.page == "sales":
    # Load data for Bobby Aquatic 1 and 2
    folder_path_1 = "./data/Bobby Aquatic 1"
    sheet_name_1 = 'Penjualan'
    penjualan_data_1 = load_data_1(folder_path_1, sheet_name_1)

    folder_path_2 = "./data/Bobby Aquatic 2"
    sheet_name_2 = 'Penjualan'
    penjualan_data_2 = load_data_2(folder_path_2, sheet_name_2)

    # Forecast data for Bobby Aquatic 1 and 2
    daily_profit_1, hw_forecast_future_1, best_seasonal_period_1, best_mae_1 = forecast_profit_1(penjualan_data_1)
    daily_profit_2, hw_forecast_future_2, best_seasonal_period_2, best_mae_2 = forecast_profit_2(penjualan_data_2)

    # Tabs for Bobby Aquatic 1 and 2
    tab1, tab2 = st.tabs(["Bobby Aquatic 1", "Bobby Aquatic 2"])

    # Bobby Aquatic 1 dashboard
    with tab1:
        st.header("Dashboard Cabang 1: Bobby Aquatic 1 - Peramalan Penjualan")
        show_dashboard_1(daily_profit_1, hw_forecast_future_1, key_suffix='cabang1')

    # Bobby Aquatic 2 dashboard
    with tab2:
        st.header("Dashboard Cabang 2: Bobby Aquatic 2 - Peramalan Penjualan")
        show_dashboard_2(daily_profit_2, hw_forecast_future_2, key_suffix='cabang2')

elif st.session_state.page == "product":
    # Load and show clustering data for Bobby Aquatic 1 and 2
    folder_path_1 = "./data/Bobby Aquatic 1"
    sheet_name_1 = 'Penjualan'
    cluster_data_1 = load_cluster_data_1(folder_path_1, sheet_name_1)

    folder_path_2 = "./data/Bobby Aquatic 2"
    sheet_name_2 = 'Penjualan'
    cluster_data_2 = load_cluster_data_2(folder_path_2, sheet_name_2)

    # Tabs for Bobby Aquatic 1 and 2
    tab1, tab2 = st.tabs(["Bobby Aquatic 1", "Bobby Aquatic 2"])

    # Bobby Aquatic 1 clustering
    with tab1:
        st.header("Klaster Produk untuk Bobby Aquatic 1")
        # Filter di samping header
        product_filter_1 = st.selectbox("Filter Produk", options=["Semua", "Filter 1", "Filter 2"], key='filter1')
        show_cluster_dashboard_1(cluster_data_1, key_suffix='cabang1')

    # Bobby Aquatic 2 clustering
    with tab2:
        st.header("Klaster Produk untuk Bobby Aquatic 2")
        # Filter di samping header
        product_filter_2 = st.selectbox("Filter Produk", options=["Semua", "Filter 1", "Filter 2"], key='filter2')
        show_cluster_dashboard_2(cluster_data_2, key_suffix='cabang2')
