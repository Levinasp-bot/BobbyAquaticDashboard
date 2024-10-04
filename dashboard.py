import streamlit as st
from sales_forecast1 import load_all_excel_files as load_data_1, forecast_profit as forecast_profit_1, show_dashboard as show_dashboard_1
from sales_forecast2 import load_all_excel_files as load_data_2, forecast_profit as forecast_profit_2, show_dashboard as show_dashboard_2
from product_clustering import load_all_excel_files as load_cluster_data, show_dashboard as show_cluster_dashboard

# Mengatur layout agar wide mode
st.set_page_config(layout="wide")

# URL handling to switch between Sales and Product
query_params = st.experimental_get_query_params()
page = query_params.get("page", ["sales"])[0]  # Default to 'sales' if no page is specified

# CSS untuk membuat tombol sidebar, menandai tombol aktif
st.markdown(
    f"""
    <style>
    .custom-button {{
        background-color: #E0E0E0;
        color: black;
        font-weight: bold;
        text-align: center;
        padding: 10px;
        border-radius: 5px;
        margin-bottom: 10px;
    }}
    .custom-button:hover {{
        background-color: #D3D3D3;
        color: black;
        cursor: pointer;
    }}
    .custom-button-active {{
        background-color: #4285f4;
        color: white;
        font-weight: bold;
        text-align: center;
        padding: 10px;
        border-radius: 5px;
        margin-bottom: 10px;
    }}
    </style>
    """,
    unsafe_allow_html=True
)

# Menentukan tombol mana yang aktif berdasarkan halaman yang dipilih
sales_class = "custom-button-active" if page == "sales" else "custom-button"
product_class = "custom-button-active" if page == "product" else "custom-button"

# Sidebar buttons (tombol seperti gambar dengan warna berbeda berdasarkan status aktif)
st.sidebar.markdown(f'<div class="{sales_class}"><a href="?page=sales" style="text-decoration: none; color: white;">Sales</a></div>', unsafe_allow_html=True)
st.sidebar.markdown(f'<div class="{product_class}"><a href="?page=product" style="text-decoration: none; color: white;">Product</a></div>', unsafe_allow_html=True)

# Load data for Bobby Aquatic 1 and 2
folder_path_1 = "./data/Bobby Aquatic 1"
sheet_name_1 = 'Penjualan'
penjualan_data_1 = load_data_1(folder_path_1, sheet_name_1)

folder_path_2 = "./data/Bobby Aquatic 2"
sheet_name_2 = 'Penjualan'
penjualan_data_2 = load_data_2(folder_path_2, sheet_name_2)

# Menampilkan konten berdasarkan halaman yang dipilih
if page == "sales":
    # Forecast data for Bobby Aquatic 1 and 2
    daily_profit_1, hw_forecast_future_1, best_seasonal_period_1, best_mae_1 = forecast_profit_1(penjualan_data_1)
    daily_profit_2, hw_forecast_future_2, best_seasonal_period_2, best_mae_2 = forecast_profit_2(penjualan_data_2)

    # Tabs for Bobby Aquatic 1 and 2
    tab1, tab2 = st.tabs(["Bobby Aquatic 1", "Bobby Aquatic 2"])

    # Bobby Aquatic 1 dashboard
    with tab1:
        st.header("Dashboard Cabang 1: Bobby Aquatic 1 - Sales Forecast")
        show_dashboard_1(daily_profit_1, hw_forecast_future_1, key_suffix='cabang1')

    # Bobby Aquatic 2 dashboard
    with tab2:
        st.header("Dashboard Cabang 2: Bobby Aquatic 2 - Sales Forecast")
        show_dashboard_2(daily_profit_2, hw_forecast_future_2, key_suffix='cabang2')

elif page == "product":
    # Load and show clustering data for Bobby Aquatic 1 and 2
    cluster_data_1 = load_cluster_data(folder_path_1, sheet_name_1)
    cluster_data_2 = load_cluster_data(folder_path_2, sheet_name_2)

    # Tabs for Bobby Aquatic 1 and 2
    tab1, tab2 = st.tabs(["Bobby Aquatic 1", "Bobby Aquatic 2"])

    # Bobby Aquatic 1 clustering
    with tab1:
        st.header("Product Clustering for Bobby Aquatic 1")
        show_cluster_dashboard(cluster_data_1, key_suffix='cabang1')

    # Bobby Aquatic 2 clustering
    with tab2:
        st.header("Product Clustering for Bobby Aquatic 2")
        show_cluster_dashboard(cluster_data_2, key_suffix='cabang2')
