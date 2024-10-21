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
def forecast_profit(data, seasonal_period=50, forecast_horizon=50):
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

def show_dashboard(daily_profit_1=None, daily_profit_2=None, hw_forecast_future=None, forecast_horizon=50, key_suffix=''):
    col1, col2 = st.columns([1, 3])

    with col1:
        # Jika kedua cabang dipilih, hitung metrik dari gabungan dua cabang
        if daily_profit_1 is not None and daily_profit_2 is not None:
            total_profit_1 = daily_profit_1['LABA'].sum()
            total_profit_2 = daily_profit_2['LABA'].sum()
            last_week_profit = (total_profit_1 + total_profit_2) / 2
        elif daily_profit_1 is not None:
            last_week_profit = daily_profit_1['LABA'].iloc[-1]
        elif daily_profit_2 is not None:
            last_week_profit = daily_profit_2['LABA'].iloc[-1]

        predicted_profit_next_week = hw_forecast_future.iloc[0]
        profit_change_percentage = ((predicted_profit_next_week - last_week_profit) / last_week_profit) * 100 if last_week_profit else 0

        total_profit_last_week = last_week_profit * 7

        arrow = "🡅" if profit_change_percentage > 0 else "🡇"
        color = "green" if profit_change_percentage > 0 else "red"

        st.markdown(f"""
            <div style="border: 2px solid #dcdcdc; padding: 10px; margin-bottom: 10px; border-radius: 5px; text-align: center;">
                <span style="font-size: 14px;">Total Laba Minggu Ini</span><br>
                <span style="font-size: 32px; font-weight: bold;">{total_profit_last_week:,.2f}</span>
            </div>
            <div style="border: 2px solid #dcdcdc; padding: 10px; margin-bottom: 10px; border-radius: 5px; text-align: center;">
                <span style="font-size: 14px;">Rata - rata Laba Harian Minggu Ini</span><br>
                <span style="font-size: 32px; font-weight: bold;">{last_week_profit:,.2f}</span>
            </div>
            <div style="border: 2px solid #dcdcdc; padding: 10px; margin-bottom: 10px; border-radius: 5px; text-align: center;">
                <span style="font-size: 14px;">Prediksi Rata - rata Laba Harian Minggu Depan</span><br>
                <span style="font-size: 32px; font-weight: bold;">{predicted_profit_next_week:,.2f}</span>
                <br><span style='color:{color}; font-size:24px;'>{arrow} {profit_change_percentage:.2f}%</span>
            </div>
        """, unsafe_allow_html=True)

    with col2:
        st.subheader('Data Historis dan Prediksi Rata - rata Laba Mingguan')

        fig = go.Figure()

        # Jika cabang 1 dipilih, tampilkan datanya di line chart
        if daily_profit_1 is not None:
            fig.add_trace(go.Scatter(x=daily_profit_1.index, y=daily_profit_1['LABA'], mode='lines', name='Cabang 1'))

        # Jika cabang 2 dipilih, tampilkan datanya di line chart dengan warna berbeda
        if daily_profit_2 is not None:
            fig.add_trace(go.Scatter(x=daily_profit_2.index, y=daily_profit_2['LABA'], mode='lines', name='Cabang 2', line=dict(color='orange')))

        # Jika forecast ada, tambahkan prediksi ke chart
        if hw_forecast_future is not None:
            forecast_dates = pd.date_range(start=daily_profit_1.index[-1] if daily_profit_1 is not None else daily_profit_2.index[-1], periods=forecast_horizon + 1, freq='W')
            fig.add_trace(go.Scatter(x=forecast_dates, y=hw_forecast_future, mode='lines', name='Prediksi Masa Depan', line=dict(dash='dash')))

        fig.update_layout(
            xaxis_title='Tanggal',
            yaxis_title='Laba',
            hovermode='x',
            margin=dict(t=18),
            height=350
        )

        st.plotly_chart(fig)

