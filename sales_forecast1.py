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

def forecast_profit(data, category, seasonal_period=13, forecast_horizon=13):
    # Filter data by category
    category_data = data[data['KATEGORI'] == category].copy()
    category_data['TANGGAL'] = pd.to_datetime(category_data['TANGGAL'])
    
    # Set index by Tanggal, and handle missing data
    category_data.set_index('TANGGAL', inplace=True)
    category_data = category_data.resample('W').mean().interpolate()
    
    # Train the Holt-Winters model and forecast
    train_size = int(len(category_data) * 0.9)
    train = category_data[:train_size]
    
    if len(train) < seasonal_period:
        st.warning(f"Not enough data for category '{category}' to perform forecasting.")
        return None, None

    hw_model = ExponentialSmoothing(train['LABA'], trend='add', seasonal='mul', seasonal_periods=seasonal_period).fit()
    hw_forecast_future = hw_model.forecast(forecast_horizon)
    
    return category_data, hw_forecast_future

def show_dashboard(daily_profit, categories, forecast_horizon=13, key_suffix=''):
    col1, col2 = st.columns([1, 3])

    # Year filter
    selected_years = st.multiselect(
        "Pilih Tahun",
        daily_profit['TANGGAL'].dt.year.unique(),
        key=f"multiselect_{key_suffix}",
        help="Pilih tahun yang ingin ditampilkan"
    )

    # Category filter
    available_categories = daily_profit['KATEGORI'].unique()

    selected_categories = st.multiselect(
        "Pilih Kategori Produk",
        available_categories,
        default=available_categories,
        key=f"category_select_{key_suffix}",
        help="Pilih kategori produk yang ingin ditampilkan"
    )

    if not selected_categories:
        selected_categories = available_categories

    # Filter data by selected years
    filtered_profit = daily_profit[daily_profit['TANGGAL'].dt.year.isin(selected_years)]

    fig = go.Figure()

    for category in selected_categories:
        # Forecast for each selected category
        category_data, hw_forecast_future = forecast_profit(filtered_profit, category, forecast_horizon=forecast_horizon)

        if category_data is None:
            continue

        # Historical data plot
        fig.add_trace(go.Scatter(
            x=category_data.index, 
            y=category_data['LABA'], 
            mode='lines', 
            name=f'Data Historis - {category}'
        ))

        # Forecast data plot
        last_actual_date = category_data.index.max()
        forecast_dates = pd.date_range(start=last_actual_date, periods=forecast_horizon + 1, freq='W')

        combined_forecast = pd.concat([category_data.iloc[[-1]]['LABA'], hw_forecast_future])
        fig.add_trace(go.Scatter(
            x=forecast_dates, 
            y=combined_forecast, 
            mode='lines', 
            name=f'Prediksi - {category}',
            line=dict(dash='dash')
        ))

    # Update layout
    fig.update_layout(
        title='Data Historis dan Prediksi Laba per Kategori',
        xaxis_title='Tanggal',
        yaxis_title='Laba',
        hovermode='x'
    )

    with col2:
        st.plotly_chart(fig)

    with col1:
        if len(selected_categories) == 1:
            category = selected_categories[0]
            last_week_profit = category_data['LABA'].iloc[-1]
            predicted_profit_next_week = hw_forecast_future.iloc[0]
            profit_change_percentage = ((predicted_profit_next_week - last_week_profit) / last_week_profit) * 100 if last_week_profit else 0

            arrow = "🡅" if profit_change_percentage > 0 else "🡇"
            color = "green" if profit_change_percentage > 0 else "red"

            st.markdown(f"""
                <div class='boxed'>
                    <span class="profit-label">Laba Minggu Terakhir - {category}</span><br>
                    <span class="profit-value">{last_week_profit:,.2f}</span>
                </div>
                <div class='boxed'>
                    <span class="profit-label">Prediksi Laba Minggu Depan - {category}</span><br>
                    <span class="profit-value">{predicted_profit_next_week:,.2f}</span>
                    <br><span style='color:{color}; font-size:24px;'>{arrow} {profit_change_percentage:.2f}%</span>
                </div>
            """, unsafe_allow_html=True)
        else:
            st.write("Pilih satu kategori untuk melihat perubahan laba.")
