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
    daily_profit = data[['TANGGAL', 'LABA', 'KATEGORI']].copy()
    daily_profit['TANGGAL'] = pd.to_datetime(daily_profit['TANGGAL'])
    daily_profit = daily_profit.groupby(['TANGGAL', 'KATEGORI']).sum().reset_index()

    daily_profit.set_index('TANGGAL', inplace=True)
    daily_profit = daily_profit.groupby('KATEGORI').resample('W').mean().interpolate().reset_index()

    train_size = int(len(daily_profit) * 0.9)
    train, test = daily_profit[:train_size], daily_profit[train_size:]

    hw_model = ExponentialSmoothing(train['LABA'], trend='add', seasonal='mul', seasonal_periods=seasonal_period).fit()

    hw_forecast_future = hw_model.forecast(forecast_horizon)

    return daily_profit, hw_forecast_future

def show_dashboard(daily_profit, hw_forecast_future, forecast_horizon=13, key_suffix=''):
    col1, col2 = st.columns([1, 3])

    selected_years = st.multiselect(
        "Pilih Tahun",
        daily_profit['TANGGAL'].dt.year.unique(),
        key=f"multiselect_{key_suffix}",
        help="Pilih tahun yang ingin ditampilkan"
    )

    # Adding filter for category after the year filter
    available_categories = daily_profit['KATEGORI'].unique()
    
    # Inisialisasi selected_categories terlebih dahulu
    selected_categories = st.multiselect(
        "Pilih Kategori Produk",
        available_categories,
        default=available_categories,  # Default semua kategori
        key=f"category_select_{key_suffix}",
        help="Pilih kategori produk yang ingin ditampilkan"
    )

    if not selected_categories:  # Jika tidak ada kategori dipilih, pilih semua
        selected_categories = available_categories

    # Filter data berdasarkan tahun dan kategori yang dipilih
    filtered_profit = daily_profit[
        (daily_profit['TANGGAL'].dt.year.isin(selected_years)) & 
        (daily_profit['KATEGORI'].isin(selected_categories))
    ]

    with col1:
        last_week_profit = filtered_profit['LABA'].iloc[-1]
        predicted_profit_next_week = hw_forecast_future.iloc[0]
        profit_change_percentage = ((predicted_profit_next_week - last_week_profit) / last_week_profit) * 100 if last_week_profit else 0

        arrow = "🡅" if profit_change_percentage > 0 else "🡇"
        color = "green" if profit_change_percentage > 0 else "red"

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
        fig = go.Figure()

        if selected_years:
            combined_data = filtered_profit[filtered_profit['TANGGAL'].dt.year.isin(selected_years)]
            fig.add_trace(go.Scatter(x=combined_data['TANGGAL'], y=combined_data['LABA'], mode='lines', name='Data Historis'))

        last_actual_date = filtered_profit['TANGGAL'].max()
        forecast_dates = pd.date_range(start=last_actual_date, periods=forecast_horizon + 1, freq='W')

        combined_forecast = pd.concat([filtered_profit.iloc[[-1]]['LABA'], hw_forecast_future])

        fig.add_trace(go.Scatter(x=forecast_dates, y=combined_forecast, mode='lines', name='Prediksi Masa Depan', line=dict(dash='dash')))

        fig.update_layout(
            title='Data Historis dan Prediksi Laba',
            xaxis_title='Tanggal',
            yaxis_title='Laba',
            hovermode='x'
        )

        st.plotly_chart(fig)
