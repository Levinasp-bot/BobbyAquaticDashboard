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

def show_dashboard(daily_profit_1, hw_forecast_future_1, daily_profit_2=None, hw_forecast_future_2=None, forecast_horizon=12, key_suffix=''):
    col1, col2 = st.columns([1, 3])

    # Combine profit metrics from both branches
    if daily_profit_2 is not None:
        last_week_profit = daily_profit_1['LABA'].iloc[-1] + daily_profit_2['LABA'].iloc[-1]
        predicted_profit_next_week = hw_forecast_future_1.iloc[0] + hw_forecast_future_2.iloc[0]
    else:
        last_week_profit = daily_profit_1['LABA'].iloc[-1]
        predicted_profit_next_week = hw_forecast_future_1.iloc[0]

    profit_change_percentage = ((predicted_profit_next_week - last_week_profit) / last_week_profit) * 100 if last_week_profit else 0
    total_profit_last_week = last_week_profit * 7

    arrow = "🡅" if profit_change_percentage > 0 else "🡇"
    color = "green" if profit_change_percentage > 0 else "red"

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

    with col2:
        st.subheader('Data Historis dan Prediksi Rata - rata Laba Mingguan')

        historical_years = daily_profit_1.index.year.unique()
        last_actual_date = daily_profit_1.index[-1]
        forecast_dates = pd.date_range(start=last_actual_date, periods=forecast_horizon + 1, freq='W')
        forecast_years = forecast_dates.year.unique()

        all_years = sorted(set(historical_years) | set(forecast_years))
        default_years = [2024] if 2024 in all_years else []

        selected_years = st.multiselect(
            "Pilih Tahun",
            all_years,
            default=default_years,
            key=f"multiselect_{key_suffix}",
            help="Pilih tahun yang ingin ditampilkan"
        )

        fig = go.Figure()

        # Plot historical data for branch 1
        if selected_years:
            combined_data_1 = daily_profit_1[daily_profit_1.index.year.isin(selected_years)]
            fig.add_trace(go.Scatter(x=combined_data_1.index, y=combined_data_1['LABA'], mode='lines', name='Data Historis Cabang 1'))

            # Plot forecast data for branch 1
            combined_forecast_1 = pd.concat([combined_data_1.iloc[[-1]]['LABA'], hw_forecast_future_1])
            fig.add_trace(go.Scatter(x=forecast_dates, y=combined_forecast_1, mode='lines', name='Prediksi Cabang 1', line=dict(dash='dash')))

        # Plot historical data and forecast for branch 2 (if available)
        if daily_profit_2 is not None:
            combined_data_2 = daily_profit_2[daily_profit_2.index.year.isin(selected_years)]
            fig.add_trace(go.Scatter(x=combined_data_2.index, y=combined_data_2['LABA'], mode='lines', name='Data Historis Cabang 2', line=dict(color='orange')))

            combined_forecast_2 = pd.concat([combined_data_2.iloc[[-1]]['LABA'], hw_forecast_future_2])
            fig.add_trace(go.Scatter(x=forecast_dates, y=combined_forecast_2, mode='lines', name='Prediksi Cabang 2', line=dict(dash='dash', color='orange')))

        fig.update_layout(
            xaxis_title='Tanggal',
            yaxis_title='Laba',
            hovermode='x',
            margin=dict(t=18),  # Mengurangi padding atas (t = top)
            height=350  # Mengurangi tinggi chart
        )

        st.plotly_chart(fig)
