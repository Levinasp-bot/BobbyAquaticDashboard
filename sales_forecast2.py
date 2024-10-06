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
    daily_profit = data[['TANGGAL', 'LABA']].copy()
    daily_profit['TANGGAL'] = pd.to_datetime(daily_profit['TANGGAL'])
    daily_profit = daily_profit.groupby('TANGGAL').sum()
    daily_profit = daily_profit[~daily_profit.index.duplicated(keep='first')]

    daily_profit = daily_profit.asfreq('W').interpolate()

    train_size = int(len(daily_profit) * 0.9)
    train, test = daily_profit[:train_size], daily_profit[train_size:]

    hw_model = ExponentialSmoothing(train, trend='add', seasonal='mul', seasonal_periods=seasonal_period).fit()

    hw_forecast_future = hw_model.forecast(forecast_horizon)

    return daily_profit, hw_forecast_future

def show_dashboard(daily_profit, hw_forecast_future, forecast_horizon=50, key_suffix=''):
    st.title(f"Trend Penjualan dan Prediksi Laba {key_suffix}")

    st.subheader("Filter berdasarkan Tahun")
    selected_years = st.multiselect(
        "Pilih Tahun",
        daily_profit.index.year.unique(),
        key=f"multiselect_{key_suffix}"
    )

    col1, col2 = st.columns([1, 3])

    with col1:
        last_week_profit = daily_profit['LABA'].iloc[-1]
        predicted_profit_next_week = hw_forecast_future.iloc[0]
        profit_change_percentage = ((predicted_profit_next_week - last_week_profit) / last_week_profit) * 100 if last_week_profit else 0

        info_style = """
        <div style='border: 1px solid #ccc; padding: 10px; height: 100px; display: flex; flex-direction: column; justify-content: center; align-items: center;'>
            <h2 style='margin: 0; font-size: 30px;'>%s</h2>
            <p style='font-size: 12px; margin: 0;'>%s</p>
        </div>
        """

        st.markdown(info_style % (f"{last_week_profit:,.2f}", "Laba Minggu Terakhir"), unsafe_allow_html=True)
        st.markdown(info_style % (f"{predicted_profit_next_week:,.2f}", "Prediksi Laba Minggu Depan"), unsafe_allow_html=True)
        st.markdown(info_style % (f"{profit_change_percentage:.2f}%", "Persentase Perubahan"), unsafe_allow_html=True)

    with col2:
        fig = go.Figure()

        if selected_years:
            for year in selected_years:
                filtered_data = daily_profit[daily_profit.index.year == year]
                fig.add_trace(go.Scatter(x=filtered_data.index, y=filtered_data['LABA'], mode='lines', name=f'Data Historis {year}'))

        last_actual_date = daily_profit.index[-1]
        forecast_dates = pd.date_range(start=last_actual_date, periods=forecast_horizon + 1, freq='W')[1:]

        fig.add_trace(go.Scatter(x=forecast_dates, y=hw_forecast_future, mode='lines', name='Prediksi Masa Depan', line=dict(dash='dash')))

        fig.update_layout(
            title='Data Historis dan Prediksi Laba',
            xaxis_title='Tanggal',
            yaxis_title='Laba',
            hovermode='x'
        )

        st.plotly_chart(fig)
