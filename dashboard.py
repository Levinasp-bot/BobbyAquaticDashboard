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
    # Load data for Bobby Aquatic 1 and 2
    folder_path_1 = "./data/Bobby Aquatic 1"
    sheet_name_1 = 'Penjualan'
    penjualan_data_1 = load_data_1(folder_path_1, sheet_name_1)

    folder_path_2 = "./data/Bobby Aquatic 2"
    sheet_name_2 = 'Penjualan'
    penjualan_data_2 = load_data_2(folder_path_2, sheet_name_2)

    # Forecast data for Bobby Aquatic 1 and 2
    category_1 = "default"  # Adjust this to the appropriate category if needed
    category_2 = "default"  # Adjust this to the appropriate category if needed
    daily_profit_1, hw_forecast_future_1 = forecast_profit_1(penjualan_data_1, category_1)
    daily_profit_2, hw_forecast_future_2 = forecast_profit_2(penjualan_data_2, category_2)

    # Custom CSS for centered tabs
    st.markdown("""
        <style>
            .stTab {
                display: flex;
                justify-content: center;
                margin: 20px 0;
            }
            .stTab .tab {
                padding: 10px 20px;
                border: 1px solid #ddd;
                border-radius: 4px;
                cursor: pointer;
            }
            .stTab .tab.selected {
                background-color: #007BFF;
                color: white;
                border-bottom: 1px solid transparent;
            }
        </style>
    """, unsafe_allow_html=True)

    # Tabs for Bobby Aquatic 1 and 2
    tab1, tab2 = st.tabs(["Bobby Aquatic 1", "Bobby Aquatic 2"])

    # Bobby Aquatic 1 dashboard
    with tab1:
        st.header("Dashboard Penjualan Bobby Aquatic 1")
        show_dashboard_1(daily_profit_1, hw_forecast_future_1, key_suffix='cabang1')

    # Bobby Aquatic 2 dashboard
    with tab2:
        st.header("Dashboard Penjualan Bobby Aquatic 2")
        show_dashboard_2(daily_profit_2, hw_forecast_future_2, key_suffix='cabang2')

elif st.session_state.page == "product":
    # Load and show clustering data for Bobby Aquatic 1 and 2
    folder_path_1 = "./data/Bobby Aquatic 1"
    sheet_name_1 = 'Penjualan'
    cluster_data_1 = load_cluster_data_1(folder_path_1, sheet_name_1)

    folder_path_2 = "./data/Bobby Aquatic 2"
    sheet_name_2 = 'Penjualan'
    cluster_data_2 = load_cluster_data_2(folder_path_2, sheet_name_2)

    # Get unique years from data
    years_1 = sorted(cluster_data_1['TANGGAL'].dt.year.dropna().astype(int).unique())
    years_2 = sorted(cluster_data_2['TANGGAL'].dt.year.dropna().astype(int).unique())

    # Tabs for Bobby Aquatic 1 and 2
    tab1, tab2 = st.tabs(["Bobby Aquatic 1", "Bobby Aquatic 2"])

    # Bobby Aquatic 1 clustering
    with tab1:
        st.header("Klaster Produk untuk Bobby Aquatic 1")
        selected_years_1 = st.multiselect("Pilih Tahun", options=years_1, default=years_1, key='years_1')
        if selected_years_1:
            filtered_data_1 = cluster_data_1[cluster_data_1['TANGGAL'].dt.year.isin(selected_years_1)]
        else:
            filtered_data_1 = cluster_data_1
        show_cluster_dashboard_1(filtered_data_1, key_suffix='cabang1')

    # Bobby Aquatic 2 clustering
    with tab2:
        st.header("Klaster Produk untuk Bobby Aquatic 2")
        selected_years_2 = st.multiselect("Pilih Tahun", options=years_2, default=years_2, key='years_2')
        if selected_years_2:
            filtered_data_2 = cluster_data_2[cluster_data_2['TANGGAL'].dt.year.isin(selected_years_2)]
        else:
            filtered_data_2 = cluster_data_2
        show_cluster_dashboard_2(filtered_data_2, key_suffix='cabang2')
