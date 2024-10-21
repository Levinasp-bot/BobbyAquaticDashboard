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
        # Show combined metrics if both branches are selected
        if daily_profit_1 is not None and daily_profit_2 is not None:
            combined_last_week_profit = (daily_profit_1['LABA'].iloc[-1] + daily_profit_2['LABA'].iloc[-1])
            combined_predicted_profit_next_week = (hw_forecast_future_1.iloc[0] + hw_forecast_future_2.iloc[0])
            combined_total_profit_last_week = combined_last_week_profit * 7
            combined_profit_change_percentage = ((combined_predicted_profit_next_week - combined_last_week_profit) / combined_last_week_profit) * 100 if combined_last_week_profit else 0

            combined_arrow = "🡅" if combined_profit_change_percentage > 0 else "🡇"
            combined_color = "green" if combined_profit_change_percentage > 0 else "red"

            st.markdown(f"""
                <div style="border: 2px solid #dcdcdc; padding: 10px; margin-bottom: 10px; border-radius: 5px; text-align: center;">
                    <span style="font-size: 14px;">Total Laba Minggu Ini</span><br>
                    <span style="font-size: 32px; font-weight: bold;">{combined_total_profit_last_week:,.2f}</span>
                </div>
                <div style="border: 2px solid #dcdcdc; padding: 10px; margin-bottom: 10px; border-radius: 5px; text-align: center;">
                    <span style="font-size: 14px;">Rata-rata Laba Harian Minggu Ini</span><br>
                    <span style="font-size: 32px; font-weight: bold;">{combined_last_week_profit:,.2f}</span>
                </div>
                <div style="border: 2px solid #dcdcdc; padding: 10px; margin-bottom: 10px; border-radius: 5px; text-align: center;">
                    <span style="font-size: 14px;">Prediksi Rata-rata Laba Harian Minggu Depan</span><br>
                    <span style="font-size: 32px; font-weight: bold;">{combined_predicted_profit_next_week:,.2f}</span>
                    <br><span style='color:{combined_color}; font-size:24px;'>{combined_arrow} {combined_profit_change_percentage:.2f}%</span>
                </div>
            """, unsafe_allow_html=True)
        # Show metrics for individual branches only if both branches are not selected
        elif daily_profit_1 is not None:
            # Metrics for Bobby Aquatic 1
            last_week_profit_1 = daily_profit_1['LABA'].iloc[-1]
            predicted_profit_next_week_1 = hw_forecast_future_1.iloc[0]
            total_profit_last_week_1 = last_week_profit_1 * 7
            profit_change_percentage_1 = ((predicted_profit_next_week_1 - last_week_profit_1) / last_week_profit_1) * 100 if last_week_profit_1 else 0

            arrow_1 = "🡅" if profit_change_percentage_1 > 0 else "🡇"
            color_1 = "green" if profit_change_percentage_1 > 0 else "red"

            st.markdown(f"""
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
            """, unsafe_allow_html=True)

        elif daily_profit_2 is not None:
            # Metrics for Bobby Aquatic 2
            last_week_profit_2 = daily_profit_2['LABA'].iloc[-1]
            predicted_profit_next_week_2 = hw_forecast_future_2.iloc[0]
            total_profit_last_week_2 = last_week_profit_2 * 7
            profit_change_percentage_2 = ((predicted_profit_next_week_2 - last_week_profit_2) / last_week_profit_2) * 100 if last_week_profit_2 else 0

            arrow_2 = "🡅" if profit_change_percentage_2 > 0 else "🡇"
            color_2 = "green" if profit_change_percentage_2 > 0 else "red"

            st.markdown(f"""
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
            """, unsafe_allow_html=True)

    with col2:
        st.subheader('Data Historis dan Prediksi Laba')

        # Create line chart with actual and forecasted profits
        fig = go.Figure()

        # Plot historical data for Bobby Aquatic 1
        fig.add_trace(go.Scatter(
            x=daily_profit_1.index,
            y=daily_profit_1['LABA'],
            mode='lines+markers',
            name='Laba Cabang 1',
            line=dict(color='blue')
        ))

        # Plot forecast for Bobby Aquatic 1
        forecast_dates_1 = pd.date_range(start=daily_profit_1.index[-1] + pd.Timedelta(days=1), periods=forecast_horizon)
        fig.add_trace(go.Scatter(
            x=forecast_dates_1,
            y=hw_forecast_future_1,
            mode='lines+markers',
            name='Prediksi Laba Cabang 1',
            line=dict(color='blue', dash='dash')
        ))

        # Plot historical data for Bobby Aquatic 2
        fig.add_trace(go.Scatter(
            x=daily_profit_2.index,
            y=daily_profit_2['LABA'],
            mode='lines+markers',
            name='Laba Cabang 2',
            line=dict(color='red')
        ))

        # Plot forecast for Bobby Aquatic 2
        forecast_dates_2 = pd.date_range(start=daily_profit_2.index[-1] + pd.Timedelta(days=1), periods=forecast_horizon)
        fig.add_trace(go.Scatter(
            x=forecast_dates_2,
            y=hw_forecast_future_2,
            mode='lines+markers',
            name='Prediksi Laba Cabang 2',
            line=dict(color='red', dash='dash')
        ))

        fig.update_layout(
            title='Perbandingan Laba Historis dan Prediksi',
            xaxis_title='Tanggal',
            yaxis_title='Laba',
            template='plotly_white'
        )

        st.plotly_chart(fig, use_container_width=True)

