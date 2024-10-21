import os
import pandas as pd
from statsmodels.tsa.holtwinters import ExponentialSmoothing
import streamlit as st
import plotly.graph_objects as go

@st.cache_data
def load_all_excel_files(folder_path, sheet_name):
    dataframes = []
    for file in os.listdir(folder_path):
        if file.endswith('.xlsm'):
            file_path = os.path.join(folder_path, file)
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            dataframes.append(df)
    return pd.concat(dataframes, ignore_index=True)

@st.cache_data
def forecast_profit(data, seasonal_period=13, forecast_horizon=13):
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

def show_dashboard(daily_profit_1, hw_forecast_future_1, daily_profit_2, hw_forecast_future_2, forecast_horizon=12, key_suffix=''):
    col1, col2 = st.columns([1, 3])

    with col1:
        # Gabungkan tampilan statistik untuk kedua cabang
        stats_html = ""

        # Statistics for Bobby Aquatic 1
        if daily_profit_1 is not None:
            last_week_profit_1 = daily_profit_1['LABA'].iloc[-1]
            predicted_profit_next_week_1 = hw_forecast_future_1.iloc[0]
            profit_change_percentage_1 = ((predicted_profit_next_week_1 - last_week_profit_1) / last_week_profit_1) * 100 if last_week_profit_1 else 0
            total_profit_last_week_1 = last_week_profit_1 * 7

            arrow_1 = "🡅" if profit_change_percentage_1 > 0 else "🡇"
            color_1 = "green" if profit_change_percentage_1 > 0 else "red"

            stats_html += f"""
                <div style="border: 2px solid #dcdcdc; padding: 10px; margin-bottom: 10px; border-radius: 5px; text-align: center;">
                    <span style="font-size: 14px;">Total Laba Minggu Ini Cabang 1</span><br>
                    <span style="font-size: 32px; font-weight: bold;">{total_profit_last_week_1:,.2f}</span>
                </div>
                <div style="border: 2px solid #dcdcdc; padding: 10px; margin-bottom: 10px; border-radius: 5px; text-align: center;">
                    <span style="font-size: 14px;">Rata-rata Laba Harian Minggu Ini Cabang 1</span><br>
                    <span style="font-size: 32px; font-weight: bold;">{last_week_profit_1:,.2f}</span>
                </div>
                <div style="border: 2px solid #dcdcdc; padding: 10px; margin-bottom: 10px; border-radius: 5px; text-align: center;">
                    <span style="font-size: 14px;">Prediksi Rata-rata Laba Harian Minggu Depan Cabang 1</span><br>
                    <span style="font-size: 32px; font-weight: bold;">{predicted_profit_next_week_1:,.2f}</span>
                    <br><span style='color:{color_1}; font-size:24px;'>{arrow_1} {profit_change_percentage_1:.2f}%</span>
                </div>
            """

        # Statistics for Bobby Aquatic 2
        if daily_profit_2 is not None:
            last_week_profit_2 = daily_profit_2['LABA'].iloc[-1]
            predicted_profit_next_week_2 = hw_forecast_future_2.iloc[0]
            profit_change_percentage_2 = ((predicted_profit_next_week_2 - last_week_profit_2) / last_week_profit_2) * 100 if last_week_profit_2 else 0
            total_profit_last_week_2 = last_week_profit_2 * 7

            arrow_2 = "🡅" if profit_change_percentage_2 > 0 else "🡇"
            color_2 = "green" if profit_change_percentage_2 > 0 else "red"

            stats_html += f"""
                <div style="border: 2px solid #dcdcdc; padding: 10px; margin-bottom: 10px; border-radius: 5px; text-align: center;">
                    <span style="font-size: 14px;">Total Laba Minggu Ini Cabang 2</span><br>
                    <span style="font-size: 32px; font-weight: bold;">{total_profit_last_week_2:,.2f}</span>
                </div>
                <div style="border: 2px solid #dcdcdc; padding: 10px; margin-bottom: 10px; border-radius: 5px; text-align: center;">
                    <span style="font-size: 14px;">Rata-rata Laba Harian Minggu Ini Cabang 2</span><br>
                    <span style="font-size: 32px; font-weight: bold;">{last_week_profit_2:,.2f}</span>
                </div>
                <div style="border: 2px solid #dcdcdc; padding: 10px; margin-bottom: 10px; border-radius: 5px; text-align: center;">
                    <span style="font-size: 14px;">Prediksi Rata-rata Laba Harian Minggu Depan Cabang 2</span><br>
                    <span style="font-size: 32px; font-weight: bold;">{predicted_profit_next_week_2:,.2f}</span>
                    <br><span style='color:{color_2}; font-size:24px;'>{arrow_2} {profit_change_percentage_2:.2f}%</span>
                </div>
            """

        # Tampilkan statistik gabungan
        st.markdown(stats_html, unsafe_allow_html=True)

    with col2:
        st.subheader('Data Historis dan Prediksi Rata-rata Laba Mingguan')
        # Sisa kode untuk menampilkan grafik dan data historis tetap sama
