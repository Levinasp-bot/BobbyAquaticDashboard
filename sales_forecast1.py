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

    # Resample the data weekly and interpolate missing values
    daily_profit = daily_profit.resample('W').mean().interpolate()

    train_size = int(len(daily_profit) * 0.9)
    train, test = daily_profit[:train_size], daily_profit[train_size:]

    # Fit the model on the training data
    hw_model = ExponentialSmoothing(train, trend='add', seasonal='mul', seasonal_periods=seasonal_period).fit()

    # Forecast the future values (including the test period)
    hw_forecast_future = hw_model.forecast(forecast_horizon)
    test_forecast = hw_model.forecast(len(test))  # Predict the test set
    
    # Store fitted values (train predictions)
    fitted_values = hw_model.fittedvalues

    return daily_profit, fitted_values, test, test_forecast, hw_forecast_future


def show_dashboard(daily_profit_1, fitted_values_1, test_1, test_forecast_1, hw_forecast_future_1, 
                   daily_profit_2, fitted_values_2, test_2, test_forecast_2, hw_forecast_future_2, 
                   forecast_horizon=12, key_suffix=''):

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

    with col2:
        st.subheader('Data Historis, Fitted, Test, dan Prediksi Rata-rata Laba Mingguan')

    # Pastikan daily_profit_1 ada sebelum melanjutkan
        if daily_profit_1 is not None:
            historical_years_1 = daily_profit_1.index.year.unique()  # Mendapatkan tahun-tahun historis
            selected_years_1 = st.multiselect('Pilih Tahun', options=historical_years_1, default=historical_years_1)

        # Filter data berdasarkan tahun yang dipilih
            filtered_data_1 = daily_profit_1[daily_profit_1.index.year.isin(selected_years_1)]

            last_actual_date_1 = daily_profit_1.index[-1]  # Tanggal terakhir pada data
            forecast_dates_1 = pd.date_range(start=last_actual_date_1, periods=forecast_horizon + 1, freq='W')

            fig = go.Figure()

        # Plot data historis yang difilter
            fig.add_trace(go.Scatter(x=filtered_data_1.index, y=filtered_data_1['LABA'], mode='lines', name='Data Historis Cabang 1'))

        # Plot fitted values (training predictions)
            fig.add_trace(go.Scatter(x=fitted_values_1.index, y=fitted_values_1, mode='lines', name='Fitted Values Cabang 1', line=dict(dash='dot')))

        # Plot test forecasts (predictions for test set)
            fig.add_trace(go.Scatter(x=test_1.index, y=test_forecast_1, mode='lines', name='Prediksi Data Test Cabang 1', line=dict(dash='dot')))

        # Plot future forecasts
            combined_forecast_1 = pd.concat([filtered_data_1.iloc[[-1]]['LABA'], hw_forecast_future_1])
            fig.add_trace(go.Scatter(x=forecast_dates_1, y=combined_forecast_1, mode='lines', name='Prediksi Masa Depan Cabang 1', line=dict(dash='dash')))
        
            st.plotly_chart(fig)
