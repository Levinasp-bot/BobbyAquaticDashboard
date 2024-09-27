import streamlit as st
from sales_forecast1 import load_all_excel_files as load_data_1, forecast_profit as forecast_profit_1, show_dashboard as show_forecast_dashboard_1
from sales_forecast2 import load_all_excel_files as load_data_2, forecast_profit as forecast_profit_2, show_dashboard as show_forecast_dashboard_2
from product_clustering import load_all_excel_files as load_clustering_data, show_clustering_dashboard

st.set_page_config(layout="wide")

# Bobby Aquatic 1 - Sales Forecast and Clustering
folder_path_1 = "./data/Bobby Aquatic 1"
sheet_name_1 = 'Penjualan'
penjualan_data_1 = load_data_1(folder_path_1, sheet_name_1)

# Bobby Aquatic 2 - Sales Forecast and Clustering
folder_path_2 = "./data/Bobby Aquatic 2"
sheet_name_2 = 'Penjualan'
penjualan_data_2 = load_data_2(folder_path_2, sheet_name_2)

tab1, tab2 = st.tabs(["Bobby Aquatic 1", "Bobby Aquatic 2"])

with tab1:
    st.header("Dashboard Cabang 1: Bobby Aquatic 1")
    
    # Show Sales Forecast
    st.subheader("Sales Forecast")
    show_forecast_dashboard_1(penjualan_data_1, key_suffix='cabang1')
    
    # Show Clustering
    st.subheader("Clustering Analysis")
    show_clustering_dashboard(penjualan_data_1, key_suffix='cabang1')

with tab2:
    st.header("Dashboard Cabang 2: Bobby Aquatic 2")
    
    # Show Sales Forecast
    st.subheader("Sales Forecast")
    show_forecast_dashboard_2(penjualan_data_2, key_suffix='cabang2')
    
    # Show Clustering
    st.subheader("Clustering Analysis")
    show_clustering_dashboard(penjualan_data_2, key_suffix='cabang2')
