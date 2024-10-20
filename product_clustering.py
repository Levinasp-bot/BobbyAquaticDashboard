import glob 
import os
import pandas as pd
from sklearn.preprocessing import StandardScaler
from sklearn.cluster import KMeans
import streamlit as st
import plotly.graph_objects as go
from yellowbrick.cluster import KElbowVisualizer
import numpy as np

# Supaya Streamlit tidak menggunakan cache untuk fungsi yang melakukan clustering
@st.cache_data
def load_all_excel_files(folder_path, sheet_name):
    # Memuat semua file Excel dari folder yang diberikan
    all_files = glob.glob(os.path.join(folder_path, "*.xlsm"))
    dfs = []
    for file in all_files:
        df = pd.read_excel(file, sheet_name=sheet_name)
        if 'KODE BARANG' in df.columns:
            df = df.loc[:, ~df.columns.duplicated()]
        dfs.append(df)
    return pd.concat(dfs, ignore_index=True)

def process_rfm(data):
    # Memproses data RFM
    data['TANGGAL'] = pd.to_datetime(data['TANGGAL'])
    reference_date = data['TANGGAL'].max()
    
    # Mengelompokkan data berdasarkan KODE BARANG dan KATEGORI
    rfm = data.groupby(['KODE BARANG', 'KATEGORI']).agg({
        'TANGGAL': lambda x: (reference_date - x.max()).days,  # Recency
        'NAMA BARANG': 'count',  # Frequency
        'TOTAL HR JUAL': 'sum'  # Monetary
    }).reset_index()
    
    rfm.columns = ['KODE BARANG', 'KATEGORI', 'Recency', 'Frequency', 'Monetary']
    return rfm

def categorize_rfm(rfm):
    # Menghitung quartiles untuk Recency, Frequency, dan Monetary
    recency_quartiles = rfm['Recency'].quantile([0.2, 0.4, 0.6, 0.8]).values
    frequency_quartiles = rfm['Frequency'].quantile([0.2, 0.4, 0.6, 0.8]).values
    monetary_quartiles = rfm['Monetary'].quantile([0.2, 0.4, 0.6, 0.8]).values

    # Membuat bins secara dinamis berdasarkan quartiles
    recency_bins = [0] + list(recency_quartiles) + [float('inf')]
    frequency_bins = [0] + list(frequency_quartiles) + [float('inf')]
    monetary_bins = [0] + list(monetary_quartiles) + [float('inf')]

    # Label untuk 5 kategori
    recency_labels = ['Baru Saja', 'Cukup Baru', 'Cukup Lama', 'Lama', 'Sangat Lama']
    frequency_labels = ['Sangat Jarang', 'Jarang', 'Cukup Sering', 'Sering', 'Sangat Sering']
    monetary_labels = ['Sangat Rendah', 'Rendah', 'Sedang', 'Tinggi', 'Sangat Tinggi']

    # Menetapkan kategori berdasarkan bins dan labels
    rfm['Recency_Category'] = pd.cut(rfm['Recency'], bins=recency_bins, labels=recency_labels)
    rfm['Frequency_Category'] = pd.cut(rfm['Frequency'], bins=frequency_bins, labels=frequency_labels)
    rfm['Monetary_Category'] = pd.cut(rfm['Monetary'], bins=monetary_bins, labels=monetary_labels)

    return rfm

def cluster_rfm(rfm_scaled, n_clusters):
    # Melakukan clustering menggunakan KMeans
    kmeans = KMeans(n_clusters=n_clusters, init='k-means++', random_state=1)
    kmeans.fit(rfm_scaled)
    return kmeans.labels_

def plot_interactive_pie_chart(rfm, cluster_labels, category_name, custom_legends):
    # Membuat grafik pie interaktif
    rfm['Cluster'] = cluster_labels
    cluster_counts = rfm['Cluster'].value_counts().reset_index()
    cluster_counts.columns = ['Cluster', 'Count']

    available_clusters = sorted(rfm['Cluster'].unique())
    custom_legend_mapped = {cluster: custom_legends.get(cluster, f'Cluster {cluster}') for cluster in available_clusters}

    cluster_counts['Cluster'] = cluster_counts['Cluster'].map(custom_legend_mapped)

    fig = go.Figure(data=[go.Pie(
        labels=cluster_counts['Cluster'],
        values=cluster_counts['Count'],
        hole=0.3,
        textinfo='percent+label',
        pull=[0.05] * len(cluster_counts)
    )])

    fig.update_layout(
        legend=dict(
            itemsizing='constant',
            orientation='h',
            yanchor='bottom',
            y=1.02,
            xanchor='center',
            x=0.5,
            traceorder='normal'
        )
    )

    return fig

def show_cluster_table(rfm, cluster_label, custom_label, key_suffix):
    # Menampilkan tabel cluster
    st.markdown(f"##### Daftar Produk yang {custom_label}", unsafe_allow_html=True)
    
    cluster_data = rfm[rfm['Cluster'] == cluster_label]
    st.dataframe(cluster_data, width=400, height=350, key=f"cluster_table_{cluster_label}_{key_suffix}")

def process_category(rfm_category, category_name, n_clusters, key_suffix=''):
    # Memproses kategori dan menampilkan hasil
    if rfm_category.shape[0] > 0 and n_clusters > 0:
        scaler = StandardScaler()
        rfm_scaled = scaler.fit_transform(rfm_category[['Recency', 'Frequency', 'Monetary']])
        
        cluster_labels = cluster_rfm(rfm_scaled, n_clusters)
        rfm_category['Cluster'] = cluster_labels

        # Mengategorikan RFM
        rfm_category = categorize_rfm(rfm_category)

        # Menghitung rata-rata RFM untuk setiap cluster
        average_rfm_per_cluster = rfm_category.groupby('Cluster')[['Recency', 'Frequency', 'Monetary']].mean().reset_index()
        
        # Menentukan kategori linguistik berdasarkan rentang quartile
        for col in ['Recency', 'Frequency', 'Monetary']:
            quartiles = average_rfm_per_cluster[col].quantile([0.2, 0.4, 0.6, 0.8]).values
            bins = [0] + list(quartiles) + [float('inf')]
            labels = [f'Q{i+1}' for i in range(len(quartiles) + 1)]
            average_rfm_per_cluster[f'{col}_Category'] = pd.cut(
                average_rfm_per_cluster[col],
                bins=bins,
                labels=labels
            )

        # Membuat legenda untuk setiap cluster menggunakan kategori rata-rata
        custom_legends = {
            cluster: f"Rata-rata Recency: {average_rfm_per_cluster[average_rfm_per_cluster['Cluster'] == cluster]['Recency'].values[0]:.2f} ({average_rfm_per_cluster[average_rfm_per_cluster['Cluster'] == cluster]['Recency_Category'].values[0]}) | "
                     f"Frekuensi: {average_rfm_per_cluster[average_rfm_per_cluster['Cluster'] == cluster]['Frequency'].values[0]:.2f} ({average_rfm_per_cluster[average_rfm_per_cluster['Cluster'] == cluster]['Frequency_Category'].values[0]}) | "
                     f"Monetary: {average_rfm_per_cluster[average_rfm_per_cluster['Cluster'] == cluster]['Monetary'].values[0]:.2f} ({average_rfm_per_cluster[average_rfm_per_cluster['Cluster'] == cluster]['Monetary_Category'].values[0]})"
            for cluster in sorted(rfm_category['Cluster'].unique())
        }

        col1, col2 = st.columns([1, 2])  

        with col1:
            st.markdown(f"<h4 style='font-size: 20px;'>Total {category_name} Terjual</h4>", unsafe_allow_html=True)
            st.markdown(f"<div style='border: 1px solid #d3d3d3; padding: 20px; border-radius: 5px; "
                        f"font-size: 32px; font-weight: bold; display: flex; justify-content: center; align-items: center; "
                        f"height: 100px;'>"
                        f"<strong>{rfm_category['Frequency'].sum()}</strong></div>", unsafe_allow_html=True)

        with col2:
            st.markdown("<h4 style='font-size: 20px;'>Rata - rata RFM</h4>", unsafe_allow_html=True)
            average_rfm = rfm_category[['Recency', 'Frequency', 'Monetary']].mean()
    
            st.markdown(f"<div style='border: 1px solid #d3d3d3; padding: 20px; border-radius: 5px; "
                        f"display: flex; justify-content: space-around; align-items: center; height: 100px;'>"
                        f"<div style='text-align: center;'>"
                        f"<span style='font-size: 32px; font-weight: bold;'>{average_rfm['Recency']:.2f}</span><br>Recency"
                        f"</div>"
                        f"<div style='text-align: center;'>"
                        f"<span style='font-size: 32px; font-weight: bold;'>{average_rfm['Frequency']:.2f}</span><br>Frequency"
                        f"</div>"
                        f"<div style='text-align: center;'>"
                        f"<span style='font-size: 32px; font-weight: bold;'>{average_rfm['Monetary']:.2f}</span><br>Monetary"
                        f"</div>"
                        f"</div>", unsafe_allow_html=True)

        # Menampilkan pie chart
        pie_chart = plot_interactive_pie_chart(rfm_category, cluster_labels, category_name, custom_legends)
        st.plotly_chart(pie_chart)

        # Menampilkan tabel untuk setiap cluster
        for cluster in sorted(rfm_category['Cluster'].unique()):
            show_cluster_table(rfm_category, cluster, f'Di Cluster {cluster}', key_suffix)

def get_optimal_k(data_scaled):
    model = KMeans(random_state=1)
    visualizer = KElbowVisualizer(model, k=(1, 11), timings=False)
    visualizer.fit(data_scaled)
    return visualizer.elbow_value_

def show_dashboard(data, key_suffix=''):
 
    rfm = process_rfm(data)

    rfm_ikan = rfm[rfm['KATEGORI'] == 'Ikan']
    n_clusters_ikan = get_optimal_k(StandardScaler().fit_transform(rfm_ikan[['Recency', 'Frequency', 'Monetary']]))
    process_category(rfm_ikan, 'Ikan', n_clusters_ikan, key_suffix)

    rfm_aksesoris = rfm[rfm['KATEGORI'] == 'Aksesoris']
    n_clusters_aksesoris = get_optimal_k(StandardScaler().fit_transform(rfm_aksesoris[['Recency', 'Frequency', 'Monetary']]))
    process_category(rfm_aksesoris, 'Aksesoris', n_clusters_aksesoris, key_suffix)