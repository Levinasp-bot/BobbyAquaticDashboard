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

# Define your categories (replace with actual categories)
    categories_1 = ["Ikan", "Aksesoris"]  # Replace with actual categories for Bobby Aquatic 1
    categories_2 = ["Ikan", "Aksesoris"]  # Replace with actual categories for Bobby Aquatic 2

# Initialize lists to store daily profits and forecasts
    daily_profits_1 = {}
    hw_forecasts_future_1 = {}
    daily_profits_2 = {}
    hw_forecasts_future_2 = {}

# Forecast data for Bobby Aquatic 1
    for category in categories_1:
        daily_profit, hw_forecast_future = forecast_profit_1(penjualan_data_1, category)
        daily_profits_1[category] = daily_profit
        hw_forecasts_future_1[category] = hw_forecast_future

# Forecast data for Bobby Aquatic 2
    for category in categories_2:
        daily_profit, hw_forecast_future = forecast_profit_2(penjualan_data_2, category)
        daily_profits_2[category] = daily_profit
        hw_forecasts_future_2[category] = hw_forecast_future

# Now you can show the dashboards based on the profits for each category
# Custom CSS for centered tabs
    st.markdown("""<style>
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
    </style>""", unsafe_allow_html=True)

# Tabs for Bobby Aquatic 1 and 2
    tab1, tab2 = st.tabs(["Bobby Aquatic 1", "Bobby Aquatic 2"])

# Bobby Aquatic 1 dashboard
    with tab1:
        st.header("Dashboard Penjualan Bobby Aquatic 1")
        for category in categories_1:
            st.subheader(f"Category: {category}")
            show_dashboard_1(daily_profits_1[category], hw_forecasts_future_1[category], key_suffix=category)

# Bobby Aquatic 2 dashboard
    with tab2:
        st.header("Dashboard Penjualan Bobby Aquatic 2")
        for category in categories_2:
            st.subheader(f"Category: {category}")
            show_dashboard_2(daily_profits_2[category], hw_forecasts_future_2[category], key_suffix=category)

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
