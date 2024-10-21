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

    hw_forecast_test = hw_model.forecast(len(test))

    return daily_profit, hw_forecast_future, train, test, hw_forecast_test

def show_dashboard(daily_profit_1, hw_forecast_future_1, train_1, test_1, hw_forecast_test_1, 
                   daily_profit_2, hw_forecast_future_2, train_2, test_2, hw_forecast_test_2, 
                   forecast_horizon=12, key_suffix=''):
    col1, col2 = st.columns([1, 3])

    with col2:
        st.subheader('Data Historis dan Prediksi Rata-rata Laba Mingguan')

        historical_years_1 = daily_profit_1.index.year.unique() if daily_profit_1 is not None else []
        historical_years_2 = daily_profit_2.index.year.unique() if daily_profit_2 is not None else []

        last_actual_date_1 = daily_profit_1.index[-1] if daily_profit_1 is not None else None
        last_actual_date_2 = daily_profit_2.index[-1] if daily_profit_2 is not None else None

        forecast_dates_1 = pd.date_range(start=last_actual_date_1, periods=forecast_horizon + 1, freq='W') if last_actual_date_1 is not None else None
        forecast_dates_2 = pd.date_range(start=last_actual_date_2, periods=forecast_horizon + 1, freq='W') if last_actual_date_2 is not None else None

        all_years = sorted(set(historical_years_1) | set(historical_years_2))
        default_years = [2024] if 2024 in all_years else []

        selected_years = st.multiselect(
            "Pilih Tahun",
            all_years,
            default=default_years,
            key=f"multiselect_{key_suffix}",
            help="Pilih tahun yang ingin ditampilkan"
        )

        fig = go.Figure()

        # Plot data for Bobby Aquatic 1
        if selected_years and daily_profit_1 is not None:
            combined_data_1 = daily_profit_1[daily_profit_1.index.year.isin(selected_years)]
            fig.add_trace(go.Scatter(x=combined_data_1.index, y=combined_data_1['LABA'], mode='lines', name='Data Historis Cabang 1'))

            if not combined_data_1.empty:
                combined_forecast_1 = pd.concat([combined_data_1.iloc[[-1]]['LABA'], hw_forecast_future_1])
                fig.add_trace(go.Scatter(x=forecast_dates_1, y=combined_forecast_1, mode='lines', name='Prediksi Masa Depan Cabang 1', line=dict(dash='dash')))

                # Plot test data
                fig.add_trace(go.Scatter(x=test_1.index, y=test_1['LABA'], mode='lines', name='Data Test Cabang 1', line=dict(color='red')))
                fig.add_trace(go.Scatter(x=test_1.index, y=hw_forecast_test_1, mode='lines', name='Prediksi Test Cabang 1', line=dict(dash='dot')))

        # Plot data for Bobby Aquatic 2
        if selected_years and daily_profit_2 is not None:
            combined_data_2 = daily_profit_2[daily_profit_2.index.year.isin(selected_years)]
            fig.add_trace(go.Scatter(x=combined_data_2.index, y=combined_data_2['LABA'], mode='lines', name='Data Historis Cabang 2'))

            if not combined_data_2.empty:
                combined_forecast_2 = pd.concat([combined_data_2.iloc[[-1]]['LABA'], hw_forecast_future_2])
                fig.add_trace(go.Scatter(x=forecast_dates_2, y=combined_forecast_2, mode='lines', name='Prediksi Masa Depan Cabang 2', line=dict(dash='dash')))

                # Plot test data
                fig.add_trace(go.Scatter(x=test_2.index, y=test_2['LABA'], mode='lines', name='Data Test Cabang 2', line=dict(color='red')))
                fig.add_trace(go.Scatter(x=test_2.index, y=hw_forecast_test_2, mode='lines', name='Prediksi Test Cabang 2', line=dict(dash='dot')))

        fig.update_layout(
            xaxis_title='Tanggal',
            yaxis_title='Laba',
            hovermode='x',
            margin=dict(t=18),
            height=350
        )

        st.plotly_chart(fig)