import streamlit as st
from sales_forecast1 import load_all_excel_files as load_data_1, forecast_profit as forecast_profit_1, show_dashboard as show_dashboard_1
from sales_forecast2 import load_all_excel_files as load_data_2, forecast_profit as forecast_profit_2, show_dashboard as show_dashboard_2
from product_clustering import load_all_excel_files as load_cluster_data, show_dashboard as show_cluster_dashboard

# Mengatur layout agar wide mode
st.set_page_config(layout="wide")

# Inisialisasi session state untuk halaman jika belum ada
if 'page' not in st.session_state:
    st.session_state.page = 'sales'  # Set default page

# Fungsi untuk mengganti halaman
def switch_page(page_name):
    st.session_state.page = page_name

# CSS untuk membuat tombol sidebar dan menandai tombol aktif
st.markdown(
    f"""
    <style>
    .custom-button {{
        background-color: #E0E0E0;
        color: black;
        font-weight: bold;
        text-align: center;
        padding: 10px;
        border-radius: 5px;
        margin-bottom: 10px;
    }}
    .custom-button:hover {{
        background-color: #D3D3D3;
        color: black;
        cursor: pointer;
    }}
    .custom-button-active {{
        background-color: #4285f4;
        color: white;
        font-weight: bold;
        text-align: center;
        padding: 10px;
        border-radius: 5px;
        margin-bottom: 10px;
    }}
    </style>
    """,
    unsafe_allow_html=True
)

# Menentukan tombol mana yang aktif berdasarkan halaman yang dipilih
sales_class = "custom-button-active" if st.session_state.page == "sales" else "custom-button"
product_class = "custom-button-active" if st.session_state.page == "product" else "custom-button"

# Sidebar buttons (tombol di sidebar yang menampilkan halaman)
st.sidebar.markdown(f'<div class="{sales_class}"><button onclick="window.location.reload();">{st.session_state.page=="sales"}</button><div class="{product_class}"><button onclick="window
