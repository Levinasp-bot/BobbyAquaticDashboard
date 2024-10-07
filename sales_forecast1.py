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
def forecast_profit(data, seasonal_period=13, forecast_horizon=13):
    daily_profit = data[['TANGGAL', 'LABA']].copy()
    daily_profit['TANGGAL'] = pd.to_datetime(daily_profit['TANGGAL'])
    daily_profit = daily_profit.groupby('TANGGAL').sum()
    daily_profit = daily_profit[~daily_profit.index.duplicated(keep='first')]

    daily_profit = daily_profit.resample('W').sum() / daily_profit.resample('W').count()

    train_size = int(len(daily_profit) * 0.9)
    train, test = daily_profit[:train_size], daily_profit[train_size:]

    hw_model = ExponentialSmoothing(train, trend='add', seasonal='mul', seasonal_periods=seasonal_period).fit()

    hw_forecast_future = hw_model.forecast(forecast_horizon)

    return daily_profit, hw_forecast_future

def show_dashboard(daily_profit, hw_forecast_future, forecast_horizon=13, key_suffix=''):
    col1, col2 = st.columns([1, 3])

    with col1:
        last_week_profit = daily_profit['LABA'].iloc[-1]
        predicted_profit_next_week = hw_forecast_future.iloc[0]
        profit_change_percentage = ((predicted_profit_next_week - last_week_profit) / last_week_profit) * 100 if last_week_profit else 0

        # Add arrows based on profit change
        if profit_change_percentage > 0:
            arrow = "🡅"
            color = "green"
        else:
            arrow = "🡇"
            color = "red"

        # Display with box outline and larger font for profit numbers
        st.markdown(""" 
            <style>
                .boxed {
                    border: 2px solid #dcdcdc;
                    padding: 10px;
                    margin-bottom: 10px;
                    border-radius: 5px;
                    text-align: center;
                }
                .profit-value {
                    font-size: 36px;
                    font-weight: bold;
                }
                .profit-label {
                    font-size: 14px;
                }
            </style>
        """, unsafe_allow_html=True)

        st.markdown(f"""
            <div class='boxed'>
                <span class="profit-label">Laba Minggu Terakhir</span><br>
                <span class="profit-value">{last_week_profit:,.2f}</span>
            </div>
            <div class='boxed'>
                <span class="profit-label">Prediksi Laba Minggu Depan</span><br>
                <span class="profit-value">{predicted_profit_next_week:,.2f}</span>
                <br><span style='color:{color}; font-size:24px;'>{arrow} {profit_change_percentage:.2f}%</span>
            </div>
        """, unsafe_allow_html=True)

    with col2:
        # Set default selected year to 2024
        default_years = [2024] if 2024 in daily_profit.index.year.unique() else []

        # Filter tahun masih ada, memungkinkan pengguna memilih beberapa tahun
        selected_years = st.multiselect(
            "Pilih Tahun",
            daily_profit.index.year.unique(),
            default=default_years,
            key=f"multiselect_{key_suffix}",
            help="Pilih tahun yang ingin ditampilkan"
        )

        fig = go.Figure()

        if selected_years:
            # Gabungkan data dari tahun yang dipilih menjadi satu garis
            combined_data = daily_profit[daily_profit.index.year.isin(selected_years)]
            fig.add_trace(go.Scatter(x=combined_data.index, y=combined_data['LABA'], mode='lines', name='Data Historis'))

        # Tambahkan prediksi masa depan
        last_actual_date = daily_profit.index[-1]
        forecast_dates = pd.date_range(start=last_actual_date, periods=forecast_horizon + 1, freq='W')

        # Gabungkan prediksi dengan titik terakhir dari data historis
        combined_forecast = pd.concat([daily_profit.iloc[[-1]]['LABA'], hw_forecast_future])

        fig.add_trace(go.Scatter(x=forecast_dates, y=combined_forecast, mode='lines', name='Prediksi Masa Depan', line=dict(dash='dash')))

        fig.update_layout(
            title='Data Historis dan Prediksi Laba',
            xaxis_title='Tanggal',
            yaxis_title='Laba',
            hovermode='x'
        )

        st.plotly_chart(fig)
