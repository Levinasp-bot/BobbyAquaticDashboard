import glob
import os
import pandas as pd
from sklearn.preprocessing import StandardScaler
from sklearn.cluster import KMeans
import streamlit as st
import plotly.express as px

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

# Fungsi untuk membuat pie chart
def plot_interactive_pie_chart(rfm, cluster_labels, col1, key_suffix):
    rfm['Cluster'] = cluster_labels
    cluster_counts = rfm['Cluster'].value_counts().reset_index()
    cluster_counts.columns = ['Cluster', 'Count']

    fig = px.pie(cluster_counts, values='Count', names='Cluster', title='Cluster Distribution', hole=0.3)
    fig.update_traces(textinfo='percent+label', pull=[0.05]*len(cluster_counts))

    col1.plotly_chart(fig, use_container_width=True)
    return rfm

def show_cluster_table(rfm, cluster_label, col2, key_suffix):
    col2.subheader(f"Cluster {cluster_label} Members")
    cluster_data = rfm[rfm['Cluster'] == cluster_label]
    col2.dataframe(cluster_data, key=f"cluster_table_{cluster_label}_{key_suffix}")

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
    k_aksesoris = 5  # Nilai k untuk kategori 'Aksesoris'

    def process_category(rfm_category, category_name, n_clusters, key_suffix):
        if rfm_category.shape[0] > 0:
            # Scale the RFM data
            rfm_scaled = scaler.fit_transform(rfm_category[['Recency', 'Frequency', 'Monetary']])
            
            # Cluster the data with defined k
            cluster_labels = cluster_rfm(rfm_scaled, n_clusters)
            
            # Add cluster labels to the dataframe
            rfm_category['Cluster'] = cluster_labels

            # Membuat dua kolom
            col1, col2 = st.columns(2)

            # Pie chart di kolom kiri
            col1.subheader(f"Cluster Distribution for {category_name} Visualization")
            rfm_with_clusters = plot_interactive_pie_chart(rfm_category, cluster_labels, col1, key_suffix)

            # Selectbox untuk memilih cluster di kolom kiri
            cluster_to_show = col1.selectbox(f'Select a cluster for {category_name}:', 
                                             sorted(rfm_with_clusters['Cluster'].unique()), 
                                             key=f'selectbox_{category_name}_{key_suffix}_{cluster_labels}')

            # Tabel detail di kolom kanan
            show_cluster_table(rfm_with_clusters, cluster_to_show, col2, key_suffix)
        else:
            st.error(f"Tidak ada data yang valid untuk clustering di kategori {category_name}.")

    # Process clustering for 'Ikan' and 'Aksesoris' with predefined k values
    process_category(rfm_ikan, 'Ikan', k_ikan, key_suffix='ikan')
    process_category(rfm_aksesoris, 'Aksesoris', k_aksesoris, key_suffix='aksesoris')
