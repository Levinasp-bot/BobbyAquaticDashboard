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

def show_dashboard(daily_profit_1=None, daily_profit_2=None, hw_forecast_1=None, hw_forecast_2=None, key_suffix=''):
    col1, col2 = st.columns([1, 3])

    # Combine the total, average, and forecast profits for both branches
    total_profit_last_week = 0
    last_week_profit = 0
    predicted_profit_next_week = 0

    if daily_profit_1 is not None:
        last_week_profit += daily_profit_1['LABA'].iloc[-1]
    
        # Check if the forecast object is valid before accessing it
        if hw_forecast_1 is not None:
            predicted_profit_next_week += hw_forecast_1.iloc[0]

        total_profit_last_week += last_week_profit * 7

    if daily_profit_2 is not None:
        last_week_profit += daily_profit_2['LABA'].iloc[-1]
    
    # Check if the forecast object is valid before accessing it
        if hw_forecast_2 is not None:
            predicted_profit_next_week += hw_forecast_2.iloc[0]

    total_profit_last_week += last_week_profit * 7

    profit_change_percentage = ((predicted_profit_next_week - last_week_profit) / last_week_profit) * 100 if last_week_profit else 0
    arrow = "🡅" if profit_change_percentage > 0 else "🡇"
    color = "green" if profit_change_percentage > 0 else "red"

    # Metrics display
    with col1:
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

    # Line chart for historical and forecast data
    with col2:
        st.subheader('Data Historis dan Prediksi Rata - rata Laba Mingguan')

        fig = go.Figure()

        # Plotting for each branch separately
        if daily_profit_1 is not None:
            fig.add_trace(go.Scatter(
                x=daily_profit_1.index,
                y=daily_profit_1['LABA'],
                mode='lines',
                name='Cabang 1'
            ))

        if daily_profit_2 is not None:
            fig.add_trace(go.Scatter(
                x=daily_profit_2.index,
                y=daily_profit_2['LABA'],
                mode='lines',
                name='Cabang 2',
                line=dict(color='orange')
            ))

        # Future prediction for both branches
        if hw_forecast_1 is not None:
            forecast_dates = pd.date_range(start=daily_profit_1.index[-1], periods=len(hw_forecast_1), freq='W')
            fig.add_trace(go.Scatter(
                x=forecast_dates,
                y=hw_forecast_1,
                mode='lines',
                name='Prediksi Cabang 1',
                line=dict(dash='dash')
            ))

        if hw_forecast_2 is not None:
            forecast_dates = pd.date_range(start=daily_profit_2.index[-1], periods=len(hw_forecast_2), freq='W')
            fig.add_trace(go.Scatter(
                x=forecast_dates,
                y=hw_forecast_2,
                mode='lines',
                name='Prediksi Cabang 2',
                line=dict(dash='dash', color='orange')
            ))

        fig.update_layout(
            xaxis_title='Tanggal',
            yaxis_title='Laba',
            hovermode='x',
            margin=dict(t=18),
            height=350
        )

        st.plotly_chart(fig)
