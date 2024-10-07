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

# Tabs untuk Bobby Aquatic 1 dan 2
if st.session_state.page == "sales":
    tab1, tab2 = st.tabs(["Bobby Aquatic 1", "Bobby Aquatic 2"])
    
    # Filter only for the active tab
    if tab1:
        # Load data for Bobby Aquatic 1
        folder_path_1 = "./data/Bobby Aquatic 1"
        sheet_name_1 = 'Penjualan'
        penjualan_data_1 = load_data_1(folder_path_1, sheet_name_1)

        # Get unique categories for Bobby Aquatic 1
        categories_1 = penjualan_data_1['KATEGORI'].unique() if 'KATEGORI' in penjualan_data_1.columns else ['All']

        # Sidebar filters for Bobby Aquatic 1
        with st.sidebar:
            st.subheader("Filter Penjualan Bobby Aquatic 1")
            selected_category_1 = st.selectbox("Pilih Kategori untuk Bobby Aquatic 1", options=['All'] + list(categories_1), key='category_select_cabang1')

        # Filter data based on category
        if selected_category_1 != 'All':
            filtered_penjualan_data_1 = penjualan_data_1[penjualan_data_1['KATEGORI'] == selected_category_1]
        else:
            filtered_penjualan_data_1 = penjualan_data_1

        # Forecast and display data for Bobby Aquatic 1
        daily_profit_1, hw_forecast_future_1 = forecast_profit_1(filtered_penjualan_data_1)
        st.header("Dashboard Penjualan Bobby Aquatic 1")
        show_dashboard_1(daily_profit_1, hw_forecast_future_1, key_suffix='cabang1')

    elif tab2:
        # Load data for Bobby Aquatic 2
        folder_path_2 = "./data/Bobby Aquatic 2"
        sheet_name_2 = 'Penjualan'
        penjualan_data_2 = load_data_2(folder_path_2, sheet_name_2)

        # Get unique categories for Bobby Aquatic 2
        categories_2 = penjualan_data_2['KATEGORI'].unique() if 'KATEGORI' in penjualan_data_2.columns else ['All']

        # Sidebar filters for Bobby Aquatic 2
        with st.sidebar:
            st.subheader("Filter Penjualan Bobby Aquatic 2")
            selected_category_2 = st.selectbox("Pilih Kategori untuk Bobby Aquatic 2", options=['All'] + list(categories_2), key='category_select_cabang2')

        # Filter data based on category
        if selected_category_2 != 'All':
            filtered_penjualan_data_2 = penjualan_data_2[penjualan_data_2['KATEGORI'] == selected_category_2]
        else:
            filtered_penjualan_data_2 = penjualan_data_2

        # Forecast and display data for Bobby Aquatic 2
        daily_profit_2, hw_forecast_future_2 = forecast_profit_2(filtered_penjualan_data_2)
        st.header("Dashboard Penjualan Bobby Aquatic 2")
        show_dashboard_2(daily_profit_2, hw_forecast_future_2, key_suffix='cabang2')
