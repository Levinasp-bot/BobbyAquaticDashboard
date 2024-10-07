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
    daily_profit['TANGGAL'] = pd.to_datetime(daily_profit['TANGGAL'], errors='coerce')
    daily_profit = daily_profit.dropna(subset=['TANGGAL'])  # Drop rows with NaT in TANGGAL
    daily_profit = daily_profit.groupby(['TANGGAL', 'KATEGORI']).sum().reset_index()

    # Set index by Tanggal, and handle missing data
    daily_profit.set_index('TANGGAL', inplace=True)
    daily_profit = daily_profit.groupby('KATEGORI').resample('W').mean().interpolate().reset_index()

    # Training and forecasting
    train_size = int(len(daily_profit) * 0.9)
    train, test = daily_profit[:train_size], daily_profit[train_size:]

    hw_model = ExponentialSmoothing(train['LABA'], trend='add', seasonal='mul', seasonal_periods=seasonal_period).fit()
    hw_forecast_future = hw_model.forecast(forecast_horizon)

    return daily_profit, hw_forecast_future

def show_dashboard(daily_profit, hw_forecast_future, forecast_horizon=13, key_suffix=''):
    col1, col2 = st.columns([1, 3])

    # Year filter
    selected_years = st.multiselect(
        "Pilih Tahun",
        daily_profit['TANGGAL'].dt.year.unique(),
        key=f"multiselect_{key_suffix}",
        help="Pilih tahun yang ingin ditampilkan"
    )

    # Adding filter for category after the year filter
    available_categories = daily_profit['KATEGORI'].unique()
    selected_categories = st.multiselect(
        "Pilih Kategori Produk",
        available_categories,
        default=available_categories,
        key=f"category_select_{key_suffix}",
        help="Pilih kategori produk yang ingin ditampilkan"
    )

    if not selected_categories:  # If no categories are selected, show all
        selected_categories = available_categories

    # Filter data based on selected year and category
    filtered_profit = daily_profit[
        (daily_profit['TANGGAL'].dt.year.isin(selected_years)) & 
        (daily_profit['KATEGORI'].isin(selected_categories))
    ]

    # Aggregate profit across all categories by date
    aggregated_profit = filtered_profit.groupby('TANGGAL').sum().reset_index()

    with col1:
        # Calculate profit changes
        if not aggregated_profit.empty:
            last_week_profit = aggregated_profit['LABA'].iloc[-1]
            predicted_profit_next_week = hw_forecast_future.iloc[0] if len(hw_forecast_future) > 0 else 0
            
            profit_change_percentage = (
                (predicted_profit_next_week - last_week_profit) / last_week_profit * 100 
                if last_week_profit else 0
            )

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
        else:
            st.warning("Tidak ada data untuk kategori atau tahun yang dipilih.")

    with col2:
        fig = go.Figure()

        # Add a single line for aggregated historical data
        fig.add_trace(go.Scatter(
            x=aggregated_profit['TANGGAL'], 
            y=aggregated_profit['LABA'], 
            mode='lines', 
            name='Data Historis - Semua Kategori'
        ))

        last_actual_date = aggregated_profit['TANGGAL'].max()
        if pd.isna(last_actual_date):
            st.error("Data tidak mencukupi untuk melakukan prediksi.")
            return

        forecast_dates = pd.date_range(start=last_actual_date, periods=forecast_horizon + 1, freq='W')

        # Combine the last actual data point with the forecast data
        combined_forecast = pd.concat([aggregated_profit.iloc[[-1]]['LABA'], hw_forecast_future])

        fig.add_trace(go.Scatter(
            x=forecast_dates, 
            y=combined_forecast, 
            mode='lines', 
            name='Prediksi Masa Depan', 
            line=dict(dash='dash')
        ))

        # Update chart layout
        fig.update_layout(
            title='Data Historis dan Prediksi Laba (Semua Kategori)',
            xaxis_title='Tanggal',
            yaxis_title='Laba',
            hovermode='x'
        )

        st.plotly_chart(fig)
