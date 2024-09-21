import streamlit as st
from sales_forecast1 import load_all_excel_files as load_data_1, forecast_profit as forecast_profit_1, show_dashboard as show_dashboard_1
from sales_forecast2 import load_all_excel_files as load_data_2, forecast_profit as forecast_profit_2, show_dashboard as show_dashboard_2

# Mengatur layout agar wide mode
st.set_page_config(layout="wide")

folder_path_1 = "./data/Bobby Aquatic 1"
sheet_name_1 = 'Penjualan'
penjualan_data_1 = load_data_1(folder_path_1, sheet_name_1)

folder_path_2 = "./data/Bobby Aquatic 2"
sheet_name_2 = 'Penjualan'
penjualan_data_2 = load_data_2(folder_path_2, sheet_name_2)

daily_profit_1, hw_forecast_future_1, best_seasonal_period_1, best_mae_1 = forecast_profit_1(penjualan_data_1)
daily_profit_2, hw_forecast_future_2, best_seasonal_period_2, best_mae_2 = forecast_profit_2(penjualan_data_2)

tab1, tab2 = st.tabs(["Bobby Aquatic 1", "Bobby Aquatic 2"])

with tab1:
    st.header("Dashboard Cabang 1: Bobby Aquatic 1")
    show_dashboard_1(daily_profit_1, hw_forecast_future_1, key_suffix='cabang1')

with tab2:
    st.header("Dashboard Cabang 2: Bobby Aquatic 2")
    show_dashboard_2(daily_profit_2, hw_forecast_future_2, key_suffix='cabang2')
