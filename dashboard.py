# Importing modules for different functionalities
import streamlit as st
from sales_forecast1 import load_all_excel_files as load_data_1, forecast_profit as forecast_profit_1, show_dashboard as show_dashboard_1
from sales_forecast2 import load_all_excel_files as load_data_2, forecast_profit as forecast_profit_2, show_dashboard as show_dashboard_2
from product_clustering import load_all_excel_files as load_cluster_data_1, show_dashboard as show_cluster_dashboard_1
from product_clustering2 import load_all_excel_files as load_cluster_data_2, show_dashboard as show_cluster_dashboard_2

st.set_page_config(layout="wide")

if 'page' not in st.session_state:
    st.session_state.page = 'sales'  

def switch_page(page_name):
    st.session_state.page = page_name
    
with st.sidebar:
    st.title("Dashboard Navigation")
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

        # Load data for Bobby Aquatic 1
        folder_path_1 = "./data/Bobby Aquatic 1"
        sheet_name_1 = 'Penjualan'
        penjualan_data_1 = load_data_1(folder_path_1, sheet_name_1)

        # Get unique categories from data
        categories_1 = penjualan_data_1['KATEGORI'].unique() if 'KATEGORI' in penjualan_data_1.columns else ['All']

        # Filter for Bobby Aquatic 1
        selected_category_1 = st.selectbox("Pilih Kategori untuk Bobby Aquatic 1", options=['All'] + list(categories_1), key='category_select_cabang1')

        # Filter data based on selected category
        filtered_penjualan_data_1 = penjualan_data_1 if selected_category_1 == 'All' else penjualan_data_1[penjualan_data_1['KATEGORI'] == selected_category_1]

        # Forecast data for Bobby Aquatic 1
        daily_profit_1, hw_forecast_future_1 = forecast_profit_1(filtered_penjualan_data_1)

        show_dashboard_1(daily_profit_1, hw_forecast_future_1, key_suffix='cabang1')

    # Bobby Aquatic 2 dashboard
    with tab2:
        st.header("Dashboard Penjualan Bobby Aquatic 2")

        # Load data for Bobby Aquatic 2
        folder_path_2 = "./data/Bobby Aquatic 2"
        sheet_name_2 = 'Penjualan'
        penjualan_data_2 = load_data_2(folder_path_2, sheet_name_2)

        # Get unique categories from data
        categories_2 = penjualan_data_2['KATEGORI'].unique() if 'KATEGORI' in penjualan_data_2.columns else ['All']

        # Filter for Bobby Aquatic 2
        selected_category_2 = st.selectbox("Pilih Kategori untuk Bobby Aquatic 2", options=['All'] + list(categories_2), key='category_select_cabang2')

        # Filter data based on selected category
        filtered_penjualan_data_2 = penjualan_data_2 if selected_category_2 == 'All' else penjualan_data_2[penjualan_data_2['KATEGORI'] == selected_category_2]

        # Forecast data for Bobby Aquatic 2
        daily_profit_2, hw_forecast_future_2 = forecast_profit_2(filtered_penjualan_data_2)

        show_dashboard_2(daily_profit_2, hw_forecast_future_2, key_suffix='cabang2')

elif st.session_state.page == "product":

    # Tabs for Bobby Aquatic 1 and 2
    tab1, tab2 = st.tabs(["Bobby Aquatic 1", "Bobby Aquatic 2"])

    # Bobby Aquatic 1 clustering dashboard
    with tab1:
        st.header("Klaster Produk untuk Bobby Aquatic 1")

        # Load data for Bobby Aquatic 1
        folder_path_1 = "./data/Bobby Aquatic 1"
        sheet_name_1 = 'Penjualan'
        cluster_data_1 = load_cluster_data_1(folder_path_1, sheet_name_1)

        # Get unique years and categories from data
        years_1 = sorted(cluster_data_1['TANGGAL'].dt.year.dropna().astype(int).unique())
        categories_1 = cluster_data_1['KATEGORI'].unique() if 'KATEGORI' in cluster_data_1.columns else ['All']

        # Filters for Bobby Aquatic 1
        selected_years_1 = st.multiselect("Pilih Tahun", options=years_1, default=years_1, key='years_1_cabang1')
        selected_category_1 = st.selectbox("Pilih Kategori", options=['All'] + list(categories_1), key='category_cluster_cabang1')

        # Filter data based on selected years and category
        filtered_data_1 = cluster_data_1[cluster_data_1['TANGGAL'].dt.year.isin(selected_years_1)]
        if selected_category_1 != 'All':
            filtered_data_1 = filtered_data_1[filtered_data_1['KATEGORI'] == selected_category_1]

        show_cluster_dashboard_1(filtered_data_1, key_suffix='cabang1')

    # Bobby Aquatic 2 clustering dashboard
    with tab2:
        st.header("Klaster Produk untuk Bobby Aquatic 2")

        # Load data for Bobby Aquatic 2
        folder_path_2 = "./data/Bobby Aquatic 2"
        sheet_name_2 = 'Penjualan'
        cluster_data_2 = load_cluster_data_2(folder_path_2, sheet_name_2)

        # Get unique years and categories from data
        years_2 = sorted(cluster_data_2['TANGGAL'].dt.year.dropna().astype(int).unique())
        categories_2 = cluster_data_2['KATEGORI'].unique() if 'KATEGORI' in cluster_data_2.columns else ['All']

        # Filters for Bobby Aquatic 2
        selected_years_2 = st.multiselect("Pilih Tahun", options=years_2, default=years_2, key='years_2_cabang2')
        selected_category_2 = st.selectbox("Pilih Kategori", options=['All'] + list(categories_2), key='category_cluster_cabang2')

        # Filter data based on selected years and category
        filtered_data_2 = cluster_data_2[cluster_data_2['TANGGAL'].dt.year.isin(selected_years_2)]
        if selected_category_2 != 'All':
            filtered_data_2 = filtered_data_2[filtered_data_2['KATEGORI'] == selected_category_2]

        show_cluster_dashboard_2(filtered_data_2, key_suffix='cabang2')
