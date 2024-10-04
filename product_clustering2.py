import glob
import os
import pandas as pd
from sklearn.preprocessing import StandardScaler
from sklearn.cluster import KMeans
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

@st.cache
def load_all_excel_files(folder_path, sheet_name):
    all_files = glob.glob(os.path.join(folder_path, "*.xlsm"))
    dfs = []
    for file in all_files:
        df = pd.read_excel(file, sheet_name=sheet_name)
        if 'KODE BARANG' in df.columns:
            df = df.loc[:, ~df.columns.duplicated()] 
        dfs.append(df)
    return pd.concat(dfs, ignore_index=True)

@st.cache
def process_rfm(data):
    data['TANGGAL'] = pd.to_datetime(data['TANGGAL'])
    reference_date = data['TANGGAL'].max()
    
    # Group by KODE BARANG and KATEGORI
    rfm = data.groupby(['KODE BARANG', 'KATEGORI']).agg({
        'TANGGAL': lambda x: (reference_date - x.max()).days,  # Recency
        'NAMA BARANG': 'count',  
        'TOTAL HR JUAL': 'sum'  
    }).reset_index()
    
    rfm.columns = ['KODE BARANG', 'KATEGORI', 'Recency', 'Frequency', 'Monetary']
    return rfm

@st.cache
def cluster_rfm(rfm_scaled, n_clusters):
    kmeans = KMeans(n_clusters=n_clusters, init='k-means++', random_state=1)
    kmeans.fit(rfm_scaled)
    return kmeans.labels_

# Fungsi untuk membuat pie chart interaktif
def plot_interactive_pie_chart(rfm, cluster_labels, key):
    rfm['Cluster'] = cluster_labels
    cluster_counts = rfm['Cluster'].value_counts().reset_index()
    cluster_counts.columns = ['Cluster', 'Count']

    fig = px.pie(cluster_counts, values='Count', names='Cluster', title='Cluster Distribution', hole=0.3)
    fig.update_traces(textinfo='percent+label', pull=[0.05]*len(cluster_counts))

    # Tambahkan interaktivitas pada pie chart
    fig.update_traces(hoverinfo='label+percent', 
                      hovertemplate='<b>Cluster %{label}</b>: %{percent}', 
                      textinfo='label+percent')

    # Deteksi event ketika sektor di-klik
    fig.data[0].on_click(lambda trace, points, selector: update_cluster_selection(points.point_inds[0], key))

    st.plotly_chart(fig, use_container_width=True)
    return rfm

# Fungsi untuk memperbarui cluster yang dipilih dalam session_state
def update_cluster_selection(selected_cluster, key):
    st.session_state[key] = selected_cluster

def show_cluster_table(rfm, key):
    selected_cluster = st.session_state.get(key, 0)
    st.subheader(f"Cluster {selected_cluster} Members")
    cluster_data = rfm[rfm['Cluster'] == selected_cluster]
    st.dataframe(cluster_data)

def show_dashboard(data, key_suffix=''):
    # Process RFM for the entire dataset
    rfm = process_rfm(data)

    # Filter RFM data by category
    rfm_ikan = rfm[rfm['KATEGORI'] == 'Ikan']
    rfm_aksesoris = rfm[rfm['KATEGORI'] == 'Aksesoris']

    # Initialize StandardScaler
    scaler = StandardScaler()

    # Definisikan jumlah cluster secara manual untuk tiap kategori
    k_ikan = 4  # Nilai k untuk kategori 'Ikan'
    k_aksesoris = 4  # Nilai k untuk kategori 'Aksesoris'

    def process_category(rfm_category, category_name, n_clusters, key_suffix):
        if rfm_category.shape[0] > 0:
            # Scale the RFM data
            rfm_scaled = scaler.fit_transform(rfm_category[['Recency', 'Frequency', 'Monetary']])
            
            # Cluster the data with defined k
            cluster_labels = cluster_rfm(rfm_scaled, n_clusters)
            
            # Add cluster labels to the dataframe
            rfm_category['Cluster'] = cluster_labels

            st.subheader(f"Cluster Distribution for {category_name} Visualization{key_suffix}")
            rfm_with_clusters = plot_interactive_pie_chart(rfm_category, cluster_labels, key=f'cluster_{category_name}')

            # Tampilkan tabel sesuai cluster yang di-klik di pie chart
            show_cluster_table(rfm_with_clusters, key=f'cluster_{category_name}')
        else:
            st.error(f"Tidak ada data yang valid untuk clustering di kategori {category_name}.")

    # Process clustering for 'Ikan' and 'Aksesoris' with predefined k values
    process_category(rfm_ikan, 'Ikan', k_ikan, key_suffix='ikan')
    process_category(rfm_aksesoris, 'Aksesoris', k_aksesoris, key_suffix='aksesoris')
