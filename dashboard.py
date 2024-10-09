import streamlit as st
from sales_forecast1 import load_all_excel_files as load_data_1, forecast_profit as forecast_profit_1, show_dashboard as show_dashboard_1
from sales_forecast2 import load_all_excel_files as load_data_2, forecast_profit as forecast_profit_2, show_dashboard as show_dashboard_2
from product_clustering import load_all_excel_files as load_cluster_data_1, show_dashboard as show_cluster_dashboard_1
from product_clustering2 import load_all_excel_files as load_cluster_data_2, show_dashboard as show_cluster_dashboard_2

# Set page configuration
st.set_page_config(page_title="Bobby Aquatic Dashboard", layout="wide")

# Custom CSS for styling
st.markdown("""
<style>
    .sidebar .sidebar-content {
        background-color: #2B8DA3; /* Sea blue for sidebar */
        padding: 20px;
        border-right: 2px solid #0D4F6D; /* Darker blue border */
    }
    .header {
        font-size: 1.5em;
        color: #1F4D68; /* Darker shade of blue for header */
    }
    .footer {
        font-size: 0.8em;
        color: #FFFFFF; /* White for footer text */
        text-align: center; /* Center the copyright text */
        margin-top: 20px; /* Added margin for spacing */
    }
    .stTabs .stTabs-header {
        justify-content: center; /* Center tab headers */
    }
    /* Adjusting button colors */
    .stButton>button {
        color: white; /* White text for buttons */
        background-color: #0D4F6D; /* Darker blue for buttons */
    }
    /* Style for select boxes and multiselects */
    .stSelectbox, .stMultiselect {
        background-color: #E0F7FA; /* Light blue for select boxes */
        color: #0D4F6D; /* Dark blue text */
    }
    /* Style for the filter options */
    .stSelectbox:hover, .stMultiselect:hover {
        background-color: #0D4F6D; /* Dark blue on hover */
        color: white; /* White text on hover */
    }
</style>
""", unsafe_allow_html=True)

# Initialize session state for page navigation
if 'page' not in st.session_state:
    st.session_state.page = 'sales'

def switch_page(page_name):
    st.session_state.page = page_name

# Sidebar for navigation
with st.sidebar:
    st.title("Dashboard Navigation")
    st.markdown("### Pilih Halaman:")
    if st.button('📊 Penjualan', key="sales_button"):
        switch_page("sales")
    if st.button('📦 Produk', key="product_button"):
        switch_page("product")

# Display content based on selected page
if st.session_state.page == "sales":
    st.header("📈 Dashboard Penjualan Bobby Aquatic")
    
    # Tabs for Bobby Aquatic 1 and 2
    tab1, tab2 = st.tabs(["Bobby Aquatic 1", "Bobby Aquatic 2"])

    # Bobby Aquatic 1 dashboard
    with tab1:
        st.header("Dashboard Penjualan Bobby Aquatic 1")

        # Load data for Bobby Aquatic 1
        folder_path_1 = "./data/Bobby Aquatic 1"
        sheet_name_1 = 'Penjualan'
        penjualan_data_1 = load_data_1(folder_path_1, sheet_name_1)

        # Get unique categories from data
        categories_1 = penjualan_data_1['KATEGORI'].dropna().unique() if 'KATEGORI' in penjualan_data_1.columns else ['Semua Kategori']

        # Filter for Bobby Aquatic 1
        selected_category_1 = st.selectbox("Pilih Kategori", options=['Semua Kategori'] + list(categories_1), key='category_select_cabang1')

        # Filter data based on selected category
        filtered_penjualan_data_1 = penjualan_data_1 if selected_category_1 == 'Semua Kategori' else penjualan_data_1[penjualan_data_1['KATEGORI'] == selected_category_1]

        # Forecast data for Bobby Aquatic 1
        daily_profit_1, hw_forecast_future_1 = forecast_profit_1(filtered_penjualan_data_1)

        # Show dashboard
        show_dashboard_1(daily_profit_1, hw_forecast_future_1, key_suffix='cabang1')

    # Bobby Aquatic 2 dashboard
    with tab2:
        st.header("Dashboard Penjualan Bobby Aquatic 2")

        # Load data for Bobby Aquatic 2
        folder_path_2 = "./data/Bobby Aquatic 2"
        sheet_name_2 = 'Penjualan'
        penjualan_data_2 = load_data_2(folder_path_2, sheet_name_2)

        # Get unique categories from data
        categories_2 = penjualan_data_2['KATEGORI'].dropna().unique() if 'KATEGORI' in penjualan_data_2.columns else ['Semua Kategori']

        # Filter for Bobby Aquatic 2
        selected_category_2 = st.selectbox("Pilih Kategori", options=['Semua Kategori'] + list(categories_2), key='category_select_cabang2')

        # Filter data based on selected category
        filtered_penjualan_data_2 = penjualan_data_2 if selected_category_2 == 'Semua Kategori' else penjualan_data_2[penjualan_data_2['KATEGORI'] == selected_category_2]

        # Forecast data for Bobby Aquatic 2
        daily_profit_2, hw_forecast_future_2 = forecast_profit_2(filtered_penjualan_data_2)

        # Show dashboard
        show_dashboard_2(daily_profit_2, hw_forecast_future_2, key_suffix='cabang2')

elif st.session_state.page == "product":
    st.header("🔍 Klaster Produk untuk Bobby Aquatic")

    # Tabs for Bobby Aquatic 1 and 2
    tab1, tab2 = st.tabs(["Bobby Aquatic 1", "Bobby Aquatic 2"])

    # Bobby Aquatic 1 clustering dashboard
    with tab1:
        st.header("Klaster Produk untuk Bobby Aquatic 1")

        # Load data for Bobby Aquatic 1
        folder_path_1 = "./data/Bobby Aquatic 1"
        sheet_name_1 = 'Penjualan'
        cluster_data_1 = load_cluster_data_1(folder_path_1, sheet_name_1)

        # Get unique years and categories from data
        years_1 = sorted(cluster_data_1['TANGGAL'].dt.year.dropna().astype(int).unique())
        categories_1 = cluster_data_1['KATEGORI'].unique() if 'KATEGORI' in cluster_data_1.columns else ['Semua Kategori']

        # Filters for Bobby Aquatic 1
        selected_years_1 = st.multiselect("Pilih Tahun", options=years_1, default=years_1, key='years_1_cabang1')
        selected_category_1 = st.selectbox("Pilih Kategori", options=['Semua Kategori'] + list(categories_1), key='category_cluster_cabang1')

        # Filter data based on selected years and category
        filtered_data_1 = cluster_data_1[cluster_data_1['TANGGAL'].dt.year.isin(selected_years_1)]
        if selected_category_1 != 'Semua Kategori':
            filtered_data_1 = filtered_data_1[filtered_data_1['KATEGORI'] == selected_category_1]

        # Show cluster dashboard
        show_cluster_dashboard_1(filtered_data_1, key_suffix='cabang1')

    # Bobby Aquatic 2 clustering dashboard
    with tab2:
        st.header("Klaster Produk untuk Bobby Aquatic 2")

        # Load data for Bobby Aquatic 2
        folder_path_2 = "./data/Bobby Aquatic 2"
        sheet_name_2 = 'Penjualan'
        cluster_data_2 = load_cluster_data_2(folder_path_2, sheet_name_2)

        # Get unique years and categories from data
        years_2 = sorted(cluster_data_2['TANGGAL'].dt.year.dropna().astype(int).unique())
        categories_2 = cluster_data_2['KATEGORI'].unique() if 'KATEGORI' in cluster_data_2.columns else ['Semua Kategori']

        # Filters for Bobby Aquatic 2
        selected_years_2 = st.multiselect("Pilih Tahun", options=years_2, default=years_2, key='years_2_cabang2')
        selected_category_2 = st.selectbox("Pilih Kategori", options=['Semua Kategori'] + list(categories_2), key='category_cluster_cabang2')

        # Filter data based on selected years and category
        filtered_data_2 = cluster_data_2[cluster_data_2['TANGGAL'].dt.year.isin(selected_years_2)]
        if selected_category_2 != 'Semua Kategori':
            filtered_data_2 = filtered_data_2[filtered_data_2['KATEGORI'] == selected_category_2]

        # Show cluster dashboard
        show_cluster_dashboard_2(filtered_data_2, key_suffix='cabang2')

# Footer
st.markdown("<div class='footer'>© 2024 Bobby Aquatic. All Rights Reserved.</div>", unsafe_allow_html=True)
