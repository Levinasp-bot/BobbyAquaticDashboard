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

    # Maintain the original layout for metrics (no changes)
    with col1:
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

    # Plotting section
    with col2:
        st.subheader('Data Historis dan Prediksi Rata-rata Laba Mingguan')

        fig = go.Figure()

        # Train data for branch 1
        if daily_profit_1 is not None:
            fig.add_trace(go.Scatter(x=daily_profit_1.index[:len(hw_forecast_future_1)], 
                                     y=daily_profit_1['LABA'].iloc[:len(hw_forecast_future_1)], 
                                     mode='lines', name='Train Cabang 1'))

            # Test data and forecast for branch 1
            fig.add_trace(go.Scatter(x=daily_profit_1.index[len(hw_forecast_future_1):], 
                                     y=daily_profit_1['LABA'].iloc[len(hw_forecast_future_1):], 
                                     mode='lines', name='Test Cabang 1'))
            fig.add_trace(go.Scatter(x=hw_forecast_future_1.index, 
                                     y=hw_forecast_future_1, 
                                     mode='lines', name='Prediksi Cabang 1'))

        # Train data for branch 2
        if daily_profit_2 is not None:
            fig.add_trace(go.Scatter(x=daily_profit_2.index[:len(hw_forecast_future_2)], 
                                     y=daily_profit_2['LABA'].iloc[:len(hw_forecast_future_2)], 
                                     mode='lines', name='Train Cabang 2'))

            # Test data and forecast for branch 2
            fig.add_trace(go.Scatter(x=daily_profit_2.index[len(hw_forecast_future_2):], 
                                     y=daily_profit_2['LABA'].iloc[len(hw_forecast_future_2):], 
                                     mode='lines', name='Test Cabang 2'))
            fig.add_trace(go.Scatter(x=hw_forecast_future_2.index, 
                                     y=hw_forecast_future_2, 
                                     mode='lines', name='Prediksi Cabang 2'))

        fig.update_layout(title="Prediksi dan Data Aktual Penjualan Cabang",
                          xaxis_title='Tanggal',
                          yaxis_title='Laba',
                          legend_title='Data',
                          hovermode='x unified')

        st.plotly_chart(fig)
