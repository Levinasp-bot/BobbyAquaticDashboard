import os
import pandas as pd
from statsmodels.tsa.holtwinters import ExponentialSmoothing
import streamlit as st
import plotly.graph_objects as go

# Menggunakan @st.cache_data sesuai Streamlit terbaru
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
    # Pastikan data hanya berisi tanggal dan laba yang sudah difilter
    daily_profit = data[['TANGGAL', 'LABA']].copy()
    daily_profit['TANGGAL'] = pd.to_datetime(daily_profit['TANGGAL'])
    daily_profit = daily_profit.groupby('TANGGAL').sum()
    daily_profit = daily_profit[~daily_profit.index.duplicated(keep='first')]

    daily_profit = daily_profit.resample('W').mean().interpolate()

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
        default_years = [2024] if 2024 in daily_profit.index.year.unique() else []

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

# Importing modules for different functionalities
from sales_forecast1 import load_all_excel_files as load_data_1, forecast_profit as forecast_profit_1, show_dashboard as show_dashboard_1
from sales_forecast2 import load_all_excel_files as load_data_2, forecast_profit as forecast_profit_2, show_dashboard as show_dashboard_2
from product_clustering import load_all_excel_files as load_cluster_data_1, show_dashboard as show_cluster_dashboard_1
from product_clustering2 import load_all_excel_files as load_cluster_data_2, show_dashboard as show_cluster_dashboard_2

# Set wide mode layout
st.set_page_config(layout="wide")

# Initialize session state for page if not already
if 'page' not in st.session_state:
    st.session_state.page = 'sales'  # Set default page

# Function to switch page
def switch_page(page_name):
    st.session_state.page = page_name

# Sidebar for navigation between Sales and Product
with st.sidebar:
    if st.button('Penjualan', key="sales_button"):
        switch_page("sales")
    if st.button('Produk', key="product_button"):
        switch_page("product")

# Display content based on selected page
if st.session_state.page == "sales":
    # Load data for Bobby Aquatic 1 and 2
    folder_path_1 = "./data/Bobby Aquatic 1"
    sheet_name_1 = 'Penjualan'
    penjualan_data_1 = load_data_1(folder_path_1, sheet_name_1)

    folder_path_2 = "./data/Bobby Aquatic 2"
    sheet_name_2 = 'Penjualan'
    penjualan_data_2 = load_data_2(folder_path_2, sheet_name_2)

    # Dapatkan kategori unik dari data
    categories_1 = penjualan_data_1['KATEGORI'].unique() if 'KATEGORI' in penjualan_data_1.columns else ['All']
    categories_2 = penjualan_data_2['KATEGORI'].unique() if 'KATEGORI' in penjualan_data_2.columns else ['All']

    # Forecast data untuk Bobby Aquatic 1 dan 2
    # Tambahkan filter kategori sebelum forecast
    selected_category_1 = st.selectbox("Pilih Kategori untuk Bobby Aquatic 1", options=['All'] + list(categories_1), key='category_select_cabang1')
    if selected_category_1 != 'All':
        filtered_penjualan_data_1 = penjualan_data_1[penjualan_data_1['KATEGORI'] == selected_category_1]
    else:
        filtered_penjualan_data_1 = penjualan_data_1

    daily_profit_1, hw_forecast_future_1 = forecast_profit_1(filtered_penjualan_data_1)

    selected_category_2 = st.selectbox("Pilih Kategori untuk Bobby Aquatic 2", options=['All'] + list(categories_2), key='category_select_cabang2')
    if selected_category_2 != 'All':
        filtered_penjualan_data_2 = penjualan_data_2[penjualan_data_2['KATEGORI'] == selected_category_2]
    else:
        filtered_penjualan_data_2 = penjualan_data_2

    daily_profit_2, hw_forecast_future_2 = forecast_profit_2(filtered_penjualan_data_2)

    # Custom CSS for centered tabs
    st.markdown("""
        <style>
            .stTab {
                display: flex;
                justify-content: center;
                margin: 20px 0;
            }
            .stTab .tab {
                padding: 10px 20px;
                border: 1px solid #ddd;
                border-radius: 4px;
                cursor: pointer;
            }
            .stTab .tab.selected {
                background-color: #007BFF;
                color: white;
                border-bottom: 1px solid transparent;
            }
        </style>
    """, unsafe_allow_html=True)

    # Tabs untuk Bobby Aquatic 1 dan 2
    tab1, tab2 = st.tabs(["Bobby Aquatic 1", "Bobby Aquatic 2"])

    # Bobby Aquatic 1 dashboard
    with tab1:
        st.header("Dashboard Penjualan Bobby Aquatic 1")
        show_dashboard_1(daily_profit_1, hw_forecast_future_1, key_suffix='cabang1')

    # Bobby Aquatic 2 dashboard
    with tab2:
        st.header("Dashboard Penjualan Bobby Aquatic 2")
        show_dashboard_2(daily_profit_2, hw_forecast_future_2, key_suffix='cabang2')

elif st.session_state.page == "product":
    # Load dan tampilkan data clustering untuk Bobby Aquatic 1 dan 2
    folder_path_1 = "./data/Bobby Aquatic 1"
    sheet_name_1 = 'Penjualan'
    cluster_data_1 = load_cluster_data_1(folder_path_1, sheet_name_1)
    
    folder_path_2 = "./data/Bobby Aquatic 2"
    sheet_name_2 = 'Penjualan'
    cluster_data_2 = load_cluster_data_2(folder_path_2, sheet_name_2)

    # Get unique years from data
    years_1 = sorted(cluster_data_1['TANGGAL'].dt.year.dropna().astype(int).unique())
    years_2 = sorted(cluster_data_2['TANGGAL'].dt.year.dropna().astype(int).unique())

    # Get unique categories from data
    categories_1 = cluster_data_1['KATEGORI'].unique() if 'KATEGORI' in cluster_data_1.columns else ['All']
    categories_2 = cluster_data_2['KATEGORI'].unique() if 'KATEGORI' in cluster_data_2.columns else ['All']

    # Tabs untuk Bobby Aquatic 1 dan 2
    tab1, tab2 = st.tabs(["Bobby Aquatic 1", "Bobby Aquatic 2"])

    # Bobby Aquatic 1 clustering
    with tab1:
        st.header("Klaster Produk untuk Bobby Aquatic 1")
        selected_years_1 = st.multiselect("Pilih Tahun", options=years_1, default=years_1, key='years_1_cabang1')
        selected_category_1 = st.selectbox("Pilih Kategori untuk Bobby Aquatic 1", options=['All'] + list(categories_1), key='category_cluster_cabang1')

        if selected_years_1:
            filtered_data_1 = cluster_data_1[cluster_data_1['TANGGAL'].dt.year.isin(selected_years_1)]
        else:
            filtered_data_1 = cluster_data_1

        if selected_category_1 != 'All':
            filtered_data_1 = filtered_data_1[filtered_data_1['KATEGORI'] == selected_category_1]

        show_cluster_dashboard_1(filtered_data_1, key_suffix='cabang1')

    # Bobby Aquatic 2 clustering
    with tab2:
        st.header("Klaster Produk untuk Bobby Aquatic 2")
        selected_years_2 = st.multiselect("Pilih Tahun", options=years_2, default=years_2, key='years_2_cabang2')
        selected_category_2 = st.selectbox("Pilih Kategori untuk Bobby Aquatic 2", options=['All'] + list(categories_2), key='category_cluster_cabang2')

        if selected_years_2:
            filtered_data_2 = cluster_data_2[cluster_data_2['TANGGAL'].dt.year.isin(selected_years_2)]
        else:
            filtered_data_2 = cluster_data_2

        if selected_category_2 != 'All':
            filtered_data_2 = filtered_data_2[filtered_data_2['KATEGORI'] == selected_category_2]

        show_cluster_dashboard_2(filtered_data_2, key_suffix='cabang2')
