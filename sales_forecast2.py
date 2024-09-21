import os
import pandas as pd
from statsmodels.tsa.holtwinters import ExponentialSmoothing
from sklearn.metrics import mean_absolute_error
import streamlit as st
import matplotlib.pyplot as plt

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
def forecast_profit(data, seasonal_min=7, seasonal_max=365, forecast_horizon=365):
    # Mengambil hanya kolom tanggal dan laba dari data penjualan
    daily_profit = data[['TANGGAL', 'LABA']].copy()
    daily_profit['TANGGAL'] = pd.to_datetime(daily_profit['TANGGAL'])
    daily_profit = daily_profit.groupby('TANGGAL').sum()
    daily_profit = daily_profit[~daily_profit.index.duplicated(keep='first')]
    daily_profit = daily_profit.asfreq('D').interpolate()

    # Membagi data
    train_size = int(len(daily_profit) * 0.9)
    train, test = daily_profit[:train_size], daily_profit[train_size:]

    best_mae = float('inf')
    best_forecast = None
    best_seasonal_period = None
    best_model = None

    # Loop untuk mencari seasonal period terbaik
    for seasonal_periods in range(seasonal_min, seasonal_max + 1):
        try:
            hw_model = ExponentialSmoothing(train, trend='add', seasonal='add', seasonal_periods=seasonal_periods).fit()
            hw_forecast_test = hw_model.forecast(len(test))
            mae_test = mean_absolute_error(test, hw_forecast_test)

            if mae_test < best_mae:
                best_mae = mae_test
                best_forecast = hw_forecast_test
                best_seasonal_period = seasonal_periods
                best_model = hw_model
        except Exception as e:
            print(f"Error with seasonal period {seasonal_periods}: {e}")

    # Prediksi 365 hari ke depan dengan model terbaik
    hw_forecast_future = best_model.forecast(forecast_horizon)
    
    return daily_profit, hw_forecast_future, best_seasonal_period, best_mae

def show_dashboard(daily_profit, hw_forecast_future, forecast_horizon=365, key_suffix=''):
    st.title(f"Dashboard Prediksi Laba Bobby Aquatic {key_suffix}")

    st.subheader("Filter berdasarkan Tahun")
    selected_years = st.multiselect(
        "Pilih Tahun",
        daily_profit.index.year.unique(),
        key=f"multiselect_{key_suffix}"
    )

    if selected_years:
        fig, ax = plt.subplots(figsize=(12, 6))

        for year in selected_years:
            filtered_data = daily_profit[daily_profit.index.year == year]
            filtered_data['LABA'].plot(ax=ax, label=f'Data Historis {year}', color='blue')

        if max(selected_years) == daily_profit.index.year.max():
            last_date = daily_profit.index[-1]
            forecast_dates = pd.date_range(start=last_date, periods=forecast_horizon + 1)[1:]
            pd.Series(hw_forecast_future, index=forecast_dates).plot(ax=ax, label='Prediksi Masa Depan', linestyle='--', color='purple')
            chart_title = f'Data Historis dan Prediksi Laba - Tahun {", ".join(map(str, selected_years))}'
        else:
            chart_title = f'Data Historis Laba - Tahun {", ".join(map(str, selected_years))}'

        ax.set_title(chart_title)
        ax.set_xlabel('Tanggal')
        ax.set_ylabel('Laba')
        ax.legend()
        st.pyplot(fig)
    else:
        st.write("Silakan pilih minimal satu tahun untuk melihat data historis.")
