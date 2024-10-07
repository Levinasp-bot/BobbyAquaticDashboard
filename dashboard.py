# Importing modules for different functionalities
import streamlit as st
from sales_forecast1 import load_all_excel_files as load_data_1, forecast_profit as forecast_profit_1, show_dashboard as show_dashboard_1
from sales_forecast2 import load_all_excel_files as load_data_2, forecast_profit as forecast_profit_2, show_dashboard as show_dashboard_2
from product_clustering import load_all_excel_files as load_cluster_data_1, show_dashboard as show_cluster_dashboard_1
from product_clustering2 import load_all_excel_files as load_cluster_data_2, show_dashboard as show_cluster_dashboard_2

# Set wide mode layout
st.set_page_config(layout="wide")

# Initialize session state for page if not already
if 'page' not in st.session_state:
    st.session_state.page = 'sales'  # Set default page

# Function to switch page
def switch_page(page_name):
    st.session_state.page = page_name

# Sidebar for navigation between Sales and Product
with st.sidebar:
    if st.button('Penjualan', key="sales_button"):
        switch_page("sales")
    if st.button('Produk', key="product_button"):
        switch_page("product")

# Display content based on selected page
if st.session_state.page == "sales":
    # Tabs for Bobby Aquatic 1 and 2
    tab1, tab2 = st.tabs(["Bobby Aquatic 1", "Bobby Aquatic 2"])

    # Bobby Aquatic 1 dashboard
    with tab1:
        st.header("Dashboard Penjualan Bobby Aquatic 1")
        
        # Load data and set filters
        folder_path_1 = "./data/Bobby Aquatic 1"
        sheet_name_1 = 'Penjualan'
        penjualan_data_1 = load_data_1(folder_path_1, sheet_name_1)
        categories_1 = penjualan_data_1['KATEGORI'].unique() if 'KATEGORI' in penjualan_data_1.columns else ['All']

        selected_category_1 = st.selectbox("Pilih Kategori untuk Bobby Aquatic 1", options=['All'] + list(categories_1), key='category_select_cabang1')
        filtered_penjualan_data_1 = penjualan_data_1[penjualan_data_1['KATEGORI'] == selected_category_1] if selected_category_1 != 'All' else penjualan_data_1

        # Forecast and display dashboard
        daily_profit_1, hw_forecast_future_1 = forecast_profit_1(filtered_penjualan_data_1)
        show_dashboard_1(daily_profit_1, hw_forecast_future_1, key_suffix='cabang1')

    # Bobby Aquatic 2 dashboard
    with tab2:
        st.header("Dashboard Penjualan Bobby Aquatic 2")
        
        # Load data and set filters
        folder_path_2 = "./data/Bobby Aquatic 2"
        sheet_name_2 = 'Penjualan'
        penjualan_data_2 = load_data_2(folder_path_2, sheet_name_2)
        categories_2 = penjualan_data_2['KATEGORI'].unique() if 'KATEGORI' in penjualan_data_2.columns else ['All']

        selected_category_2 = st.selectbox("Pilih Kategori untuk Bobby Aquatic 2", options=['All'] + list(categories_2), key='category_select_cabang2')
        filtered_penjualan_data_2 = penjualan_data_2[penjualan_data_2['KATEGORI'] == selected_category_2] if selected_category_2 != 'All' else penjualan_data_2

        # Forecast and display dashboard
        daily_profit_2, hw_forecast_future_2 = forecast_profit_2(filtered_penjualan_data_2)
        show_dashboard_2(daily_profit_2, hw_forecast_future_2, key_suffix='cabang2')

elif st.session_state.page == "product":
    # Tabs for Bobby Aquatic 1 and 2
    tab1, tab2 = st.tabs(["Bobby Aquatic 1", "Bobby Aquatic 2"])

    # Bobby Aquatic 1 clustering
    with tab1:
        st.header("Klaster Produk untuk Bobby Aquatic 1")
        
        # Load data and set filters
        folder_path_1 = "./data/Bobby Aquatic 1"
        sheet_name_1 = 'Penjualan'
        cluster_data_1 = load_cluster_data_1(folder_path_1, sheet_name_1)
        years_1 = sorted(cluster_data_1['TANGGAL'].dt.year.dropna().astype(int).unique())
        categories_1 = cluster_data_1['KATEGORI'].unique() if 'KATEGORI' in cluster_data_1.columns else ['All']

        selected_years_1 = st.multiselect("Pilih Tahun", options=years_1, default=years_1, key='years_1_cabang1')
        selected_category_1 = st.selectbox("Pilih Kategori untuk Bobby Aquatic 1", options=['All'] + list(categories_1), key='category_cluster_cabang1')

        filtered_data_1 = cluster_data_1[cluster_data_1['TANGGAL'].dt.year.isin(selected_years_1)] if selected_years_1 else cluster_data_1
        filtered_data_1 = filtered_data_1[filtered_data_1['KATEGORI'] == selected_category_1] if selected_category_1 != 'All' else filtered_data_1

        show_cluster_dashboard_1(filtered_data_1, key_suffix='cabang1')

    # Bobby Aquatic 2 clustering
    with tab2:
        st.header("Klaster Produk untuk Bobby Aquatic 2")
        
        # Load data and set filters
        folder_path_2 = "./data/Bobby Aquatic 2"
        sheet_name_2 = 'Penjualan'
        cluster_data_2 = load_cluster_data_2(folder_path_2, sheet_name_2)
        years_2 = sorted(cluster_data_2['TANGGAL'].dt.year.dropna().astype(int).unique())
        categories_2 = cluster_data_2['KATEGORI'].unique() if 'KATEGORI' in cluster_data_2.columns else ['All']

        selected_years_2 = st.multiselect("Pilih Tahun", options=years_2, default=years_2, key='years_2_cabang2')
        selected_category_2 = st.selectbox("Pilih Kategori untuk Bobby Aquatic 2", options=['All'] + list(categories_2), key='category_cluster_cabang2')

        filtered_data_2 = cluster_data_2[cluster_data_2['TANGGAL'].dt.year.isin(selected_years_2)] if selected_years_2 else cluster_data_2
        filtered_data_2 = filtered_data_2[filtered_data_2['KATEGORI'] == selected_category_2] if selected_category_2 != 'All' else filtered_data_2

        show_cluster_dashboard_2(filtered_data_2, key_suffix='cabang2')
