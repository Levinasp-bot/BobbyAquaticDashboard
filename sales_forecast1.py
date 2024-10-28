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
        # Display combined metrics if both branches have data
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

        elif daily_profit_1 is not None and daily_profit_2 is None:
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

        else:
            last_week_profit_2 = daily_profit_2['LABA'].iloc[-1]
            predicted_profit_next_week_2 = hw_forecast_future_2.iloc[0]
            total_profit_last_week_2 = last_week_profit_2 * 7
            profit_change_percentage_2 = ((predicted_profit_next_week_2 - last_week_profit_2) / last_week_profit_2) * 100 if last_week_profit_2 else 0

            arrow_1 = "🡅" if profit_change_percentage_2 > 0 else "🡇"
            color_1 = "green" if profit_change_percentage_2 > 0 else "red"

            st.markdown(f"""
                <div style="border: 2px solid #dcdcdc; padding: 10px; margin-bottom: 10px; border-radius: 5px; text-align: center;">
                    <span style="font-size: 14px;">Total Laba Minggu Ini Cabang 1</span><br>
                    <span style="font-size: 32px; font-weight: bold;">{total_profit_last_week_2:,.2f}</span>
                </div>
                <div style="border: 2px solid #dcdcdc; padding: 10px; margin-bottom: 10px; border-radius: 5px; text-align: center;">
                    <span style="font-size: 14px;">Rata-rata Laba Harian Minggu Ini Cabang 1</span><br>
                    <span style="font-size: 32px; font-weight: bold;">{last_week_profit_2:,.2f}</span>
                </div>
                <div style="border: 2px solid #dcdcdc; padding: 10px; margin-bottom: 10px; border-radius: 5px; text-align: center;">
                    <span style="font-size: 14px;">Prediksi Rata-rata Laba Harian Minggu Depan Cabang 1</span><br>
                    <span style="font-size: 32px; font-weight: bold;">{predicted_profit_next_week_2:,.2f}</span>
                    <br><span style='color:{color_1}; font-size:24px;'>{arrow_1} {profit_change_percentage_2:.2f}%</span>
                </div>
            """, unsafe_allow_html=True)

    with col2:
        st.markdown("<h3 style='font-size:20px;'>Data Historis dan Prediksi Rata-rata Laba Mingguan</h3>", unsafe_allow_html=True)

        # Gabungkan data laba harian jika kedua cabang ada, untuk menentukan daftar tahun historis yang tersedia
        combined_daily_profit = pd.concat([daily_profit_1.assign(Cabang='Cabang 1'),
                                            daily_profit_2.assign(Cabang='Cabang 2')]) if daily_profit_1 is not None and daily_profit_2 is not None else daily_profit_1 if daily_profit_1 is not None else daily_profit_2

        # Definisikan tahun historis untuk filter
        historical_years = combined_daily_profit.index.year.unique() if combined_daily_profit is not None else []
        
        default_years = [2024] if 2024 in historical_years else []
        selected_years = st.multiselect('Filter Tahun untuk Grafik', options=historical_years, default=default_years)

        if daily_profit_1 is not None and daily_profit_2 is not None:
            filtered_data = combined_daily_profit[combined_daily_profit.index.year.isin(selected_years)]
            # Filter data fitted dan prediksi untuk kedua cabang
            filtered_fitted_values_1 = fitted_values_1[fitted_values_1.index.year.isin(selected_years)]
            filtered_fitted_values_2 = fitted_values_2[fitted_values_2.index.year.isin(selected_years)]
            filtered_test_1 = test_1[test_1.index.year.isin(selected_years)]
            filtered_test_forecast_1 = test_forecast_1[test_forecast_1.index.year.isin(selected_years)]
            filtered_test_2 = test_2[test_2.index.year.isin(selected_years)]
            filtered_test_forecast_2 = test_forecast_2[test_forecast_2.index.year.isin(selected_years)]

            # Inisialisasi plot
            fig = go.Figure()
            fig.update_layout(margin=dict(t=8), height=320)

            # Plot data historis untuk kedua cabang
            for cabang in filtered_data['Cabang'].unique():
                cabang_data = filtered_data[filtered_data['Cabang'] == cabang]
                fig.add_trace(go.Scatter(x=cabang_data.index, y=cabang_data['LABA'], mode='lines', name=f'Data Historis {cabang}'))

                # Combine the last point of the historical data with the first point of fitted values
                if cabang == 'Cabang 1' and not filtered_fitted_values_1.empty:
                    # Combine the last point of the fitted values with the first point of test forecasts
                    if not filtered_test_1.empty and not filtered_test_forecast_1.empty:
                        # Shift the test forecast by one period to the right
                        shifted_test_forecast_1 = filtered_test_forecast_1.shift(1)
                        combined_test_data_1 = pd.concat([filtered_fitted_values_1.iloc[[-1]], shifted_test_forecast_1])
                        fig.add_trace(go.Scatter(x=combined_test_data_1.index, y=combined_test_data_1, mode='lines', line=dict(dash='dot', color='blue'), showlegend=False))
                        
                        # Combine the last point of the test forecast with the first point of future forecasts
                        combined_forecast_1 = pd.concat([shifted_test_forecast_1.iloc[[-1]], hw_forecast_future_1])
                        forecast_dates_1 = pd.date_range(start=cabang_data.index[-1], periods=forecast_horizon + 1, freq='W')
                        fig.add_trace(go.Scatter(x=forecast_dates_1, y=combined_forecast_1, mode='lines', name='Prediksi Laba Cabang 1', line=dict(dash='dot', color='blue')))

                elif cabang == 'Cabang 2' and not filtered_fitted_values_2.empty:
                    # Combine the last point of the fitted values with the first point of test forecasts
                    if not filtered_test_2.empty and not filtered_test_forecast_2.empty:
                        shifted_test_forecast_2 = filtered_test_forecast_2.shift(1)
                        combined_test_data_2 = pd.concat([filtered_fitted_values_2.iloc[[-1]], shifted_test_forecast_2])
                        fig.add_trace(go.Scatter(x=combined_test_data_2.index, y=combined_test_data_2, mode='lines', line=dict(dash='dot', color='orange'), showlegend=False))
                        
                        combined_forecast_2 = pd.concat([shifted_test_forecast_2.iloc[[-1]], hw_forecast_future_2])
                        forecast_dates_2 = pd.date_range(start=cabang_data.index[-1], periods=forecast_horizon + 1, freq='W')
                        fig.add_trace(go.Scatter(x=forecast_dates_2, y=combined_forecast_2, mode='lines', name='Prediksi Laba Cabang 2', line=dict(dash='dot', color='orange')))
                        
            st.plotly_chart(fig, key="plot_1")

        elif daily_profit_1 is not None:  # Only Cabang 1 is available
            filtered_data_1 = daily_profit_1[daily_profit_1.index.year.isin(selected_years)]
            filtered_fitted_values_1 = fitted_values_1[fitted_values_1.index.year.isin(selected_years)]
            filtered_test_1 = test_1[test_1.index.year.isin(selected_years)]
            filtered_test_forecast_1 = test_forecast_1[test_forecast_1.index.year.isin(selected_years)]

            fig = go.Figure()
            fig.update_layout(margin=dict(t=8), height=320)
            
            # Plot for Cabang 1
            fig.add_trace(go.Scatter(x=filtered_data_1.index, y=filtered_data_1['LABA'], mode='lines', name='Data Historis Laba Cabang 1'))

            # Fitted, Test, and Forecasts for Cabang 1
            if not filtered_fitted_values_1.empty:

                if not filtered_test_1.empty and not filtered_test_forecast_1.empty:
                    shifted_test_forecast_1 = filtered_test_forecast_1.shift(1)
                    combined_test_data_1 = pd.concat([filtered_fitted_values_1.iloc[[-1]], shifted_test_forecast_1])
                    fig.add_trace(go.Scatter(x=combined_test_data_1.index, y=combined_test_data_1, mode='lines', name='Prediksi Data Test Cabang 1', line=dict(dash='dot')))
                    
                    combined_forecast_1 = pd.concat([shifted_test_forecast_1.iloc[[-1]], hw_forecast_future_1])
                    forecast_dates_1 = pd.date_range(start=filtered_data_1.index[-1], periods=forecast_horizon + 1, freq='W')
                    fig.add_trace(go.Scatter(x=forecast_dates_1, y=combined_forecast_1, mode='lines', name='Prediksi Masa Depan Cabang 1', line=dict(dash='dot')))
            st.plotly_chart(fig, key="plot_2")

        elif daily_profit_2 is not None:  # Only Cabang 2 is available
            filtered_data_2 = daily_profit_2[daily_profit_2.index.year.isin(selected_years)]
            filtered_fitted_values_2 = fitted_values_2[fitted_values_2.index.year.isin(selected_years)]
            filtered_test_2 = test_2[test_2.index.year.isin(selected_years)]
            filtered_test_forecast_2 = test_forecast_2[test_forecast_2.index.year.isin(selected_years)]

            fig = go.Figure()
            fig.update_layout(margin=dict(t=8), height=320)
            
            # Plot for Cabang 2
            fig.add_trace(go.Scatter(x=filtered_data_2.index, y=filtered_data_2['LABA'], mode='lines', name='Data Historis Laba Cabang 2'))

            if not filtered_fitted_values_2.empty:

                if not filtered_test_2.empty and not filtered_test_forecast_2.empty:
                    shifted_test_forecast_2 = filtered_test_forecast_2.shift(1)
                    combined_test_data_2 = pd.concat([filtered_fitted_values_2.iloc[[-1]], shifted_test_forecast_2])
                    fig.add_trace(go.Scatter(x=combined_test_data_2.index, y=combined_test_data_2, mode='lines', name='Prediksi Data Test Cabang 2', line=dict(dash='dot')))
                    
                    combined_forecast_2 = pd.concat([shifted_test_forecast_2.iloc[[-1]], hw_forecast_future_2])
                    forecast_dates_2 = pd.date_range(start=filtered_data_2.index[-1], periods=forecast_horizon + 1, freq='W')
                    fig.add_trace(go.Scatter(x=forecast_dates_2, y=combined_forecast_2, mode='lines', name='Prediksi Masa Depan Cabang 2', line=dict(dash='dot')))

            st.plotly_chart(fig, key="plot_3")