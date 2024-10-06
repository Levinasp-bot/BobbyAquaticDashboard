import os
import pandas as pd
from statsmodels.tsa.holtwinters import ExponentialSmoothing
import streamlit as st
import plotly.graph_objects as go

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
def forecast_profit(data, seasonal_period=50, forecast_horizon=50):
    # Mengambil hanya kolom tanggal dan laba dari data penjualan
    daily_profit = data[['TANGGAL', 'LABA']].copy()
    daily_profit['TANGGAL'] = pd.to_datetime(daily_profit['TANGGAL'])
    daily_profit = daily_profit.groupby('TANGGAL').sum()
    daily_profit = daily_profit[~daily_profit.index.duplicated(keep='first')]

    # Resample the data weekly and interpolate missing values
    daily_profit = daily_profit.asfreq('W').interpolate()

    # Membagi data menjadi train dan test
    train_size = int(len(daily_profit) * 0.9)
    train, test = daily_profit[:train_size], daily_profit[train_size:]

    # Train the Holt-Winters model
    hw_model = ExponentialSmoothing(train, trend='add', seasonal='mul', seasonal_periods=seasonal_period).fit()

    # Forecasting mulai dari minggu setelah data historis terakhir
    hw_forecast_future = hw_model.forecast(forecast_horizon)

    return daily_profit, hw_forecast_future

def show_dashboard(daily_profit, hw_forecast_future, forecast_horizon=50, key_suffix=''):
    st.title(f"Dashboard Prediksi Laba Bobby Aquatic {key_suffix}")

    # Filter berdasarkan tahun
    st.subheader("Filter berdasarkan Tahun")
    selected_years = st.multiselect(
        "Pilih Tahun",
        daily_profit.index.year.unique(),
        key=f"multiselect_{key_suffix}"
    )

    # Prepare data for plotting
    fig = go.Figure()

    # Plot historical data
    if selected_years:
        for year in selected_years:
            filtered_data = daily_profit[daily_profit.index.year == year]
            fig.add_trace(go.Scatter(x=filtered_data.index, y=filtered_data['LABA'], mode='lines', name=f'Data Historis {year}'))

    # Ensure forecast is connected to the actual data
    last_actual_date = daily_profit.index[-1]
    forecast_dates = pd.date_range(start=last_actual_date, periods=forecast_horizon + 1, freq='W')[1:]
    
    # Plot forecast as a continuation of the actual data
    fig.add_trace(go.Scatter(x=forecast_dates, y=hw_forecast_future, mode='lines', name='Prediksi Masa Depan', line=dict(dash='dash')))

    fig.update_layout(
        title='Data Historis dan Prediksi Laba',
        xaxis_title='Tanggal',
        yaxis_title='Laba',
        hovermode='x'
    )

    st.plotly_chart(fig)