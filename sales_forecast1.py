import os
import pandas as pd
from statsmodels.tsa.holtwinters import ExponentialSmoothing
import streamlit as st
import plotly.graph_objects as go
from sales_forecast1 import load_all_excel_files as load_data_1, forecast_profit as forecast_profit_1, show_dashboard as show_dashboard_1
from sales_forecast2 import load_all_excel_files as load_data_2, forecast_profit as forecast_profit_2, show_dashboard as show_dashboard_2
from product_clustering import load_all_excel_files as load_cluster_data_1, show_dashboard as show_cluster_dashboard_1
from product_clustering2 import load_all_excel_files as load_cluster_data_2, show_dashboard as show_cluster_dashboard_2

# Set wide mode layout (this must be the first Streamlit command)
st.set_page_config(layout="wide")

@st.cache
def load_all_excel_files(folder_path, sheet_name):
    dataframes = []
    for file in os.listdir(folder_path):
        if file.endswith('.xlsm'):
            file_path = os.path.join(folder_path, file)
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            dataframes.append(df)
    return pd.concat(dataframes, ignore_index=True)

@st.cache
def forecast_profit(data, seasonal_period=13, forecast_horizon=13):
    # Ensure data contains only 'TANGGAL' and 'LABA'
    daily_profit = data[['TANGGAL', 'LABA']].copy()
    daily_profit['TANGGAL'] = pd.to_datetime(daily_profit['TANGGAL'])
    daily_profit = daily_profit.groupby('TANGGAL').sum()
    daily_profit = daily_profit[~daily_profit.index.duplicated(keep='first')]

    daily_profit = daily_profit.resample('W').mean().interpolate()

    train_size = int(len(daily_profit) * 0.9)
    train, test = daily_profit[:train_size], daily_profit[train_size:]

    hw_model = ExponentialSmoothing(train, trend='add', seasonal='mul', seasonal_periods=seasonal_period).fit()

    hw_forecast_future = hw_model.forecast(forecast_horizon)

    return daily_profit, hw_forecast_future

def show_dashboard(daily_profit, hw_forecast_future, forecast_horizon=13, key_suffix=''):
    col1, col2 = st.columns([1, 3])

    with col1:
        last_week_profit = daily_profit['LABA'].iloc[-1]
        predicted_profit_next_week = hw_forecast_future.iloc[0]
        profit_change_percentage = ((predicted_profit_next_week - last_week_profit) / last_week_profit) * 100 if last_week_profit else 0

        # Add arrows based on profit change
        if profit_change_percentage > 0:
            arrow = "🡅"
            color = "green"
        else:
            arrow = "🡇"
            color = "red"

        st.markdown(""" 
            <style>
                .boxed {
                    border: 2px solid #dcdcdc;
                    padding: 10px;
                    margin-bottom: 10px;
                    border-radius: 5px;
                    text-align: center;
                }
                .profit-value {
                    font-size: 36px;
                    font-weight: bold;
                }
                .profit-label {
                    font-size: 14px;
                }
            </style>
        """, unsafe_allow_html=True)

        st.markdown(f"""
            <div class='boxed'>
                <span class="profit-label">Laba Minggu Terakhir</span><br>
                <span class="profit-value">{last_week_profit:,.2f}</span>
            </div>
            <div class='boxed'>
                <span class="profit-label">Prediksi Laba Minggu Depan</span><br>
                <span class="profit-value">{predicted_profit_next_week:,.2f}</span>
                <br><span style='color:{color}; font-size:24px;'>{arrow} {profit_change_percentage:.2f}%</span>
            </div>
        """, unsafe_allow_html=True)

    with col2:
        default_years = [2024] if 2024 in daily_profit.index.year.unique() else []

        # Filter tahun dan kategori (misalnya produk, wilayah, dll.)
        selected_years = st.multiselect("Pilih Tahun", daily_profit.index.year.unique(), default=default_years, key=f"multiselect_{key_suffix}")
        selected_category = st.selectbox("Pilih Kategori", options=data['KATEGORI'].unique(), key=f"category_select_{key_suffix}")

        # Filter data sesuai tahun dan kategori
        filtered_data = daily_profit[daily_profit.index.year.isin(selected_years) & (data['KATEGORI'] == selected_category)]

        fig = go.Figure()

        if not filtered_data.empty:
            fig.add_trace(go.Scatter(x=filtered_data.index, y=filtered_data['LABA'], mode='lines', name='Data Historis'))

        last_actual_date = filtered_data.index[-1] if not filtered_data.empty else daily_profit.index[-1]
        forecast_dates = pd.date_range(start=last_actual_date, periods=forecast_horizon + 1, freq='W')

        combined_forecast = pd.concat([filtered_data.iloc[[-1]]['LABA'], hw_forecast_future])

        fig.add_trace(go.Scatter(x=forecast_dates, y=combined_forecast, mode='lines', name='Prediksi Masa Depan', line=dict(dash='dash')))

        fig.update_layout(
            title='Data Historis dan Prediksi Laba',
            xaxis_title='Tanggal',
            yaxis_title='Laba',
            hovermode='x'
        )

        st.plotly_chart(fig)

# Continue with the rest of your script...
