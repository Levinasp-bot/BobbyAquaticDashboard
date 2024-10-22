import glob
import os
import pandas as pd
from sklearn.preprocessing import StandardScaler
from sklearn.cluster import KMeans
import streamlit as st
import plotly.graph_objects as go
from yellowbrick.cluster import KElbowVisualizer

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
    # Ensure columns are not duplicated
    data = data.loc[:, ~data.columns.duplicated()]

    # Memproses data RFM
    data['TANGGAL'] = pd.to_datetime(data['TANGGAL'])
    reference_date = data['TANGGAL'].max()
    
    # Mengelompokkan data berdasarkan KODE BARANG dan NAMA BARANG
    rfm = data.groupby(['KODE BARANG', 'NAMA BARANG']).agg({
        'TANGGAL': lambda x: (reference_date - x.max()).days,  # Recency
        'NAMA BARANG': 'count',  # Frequency
        'TOTAL HR JUAL': 'sum'  # Monetary
    }).reset_index()
    
    rfm.columns = ['KODE BARANG', 'NAMA BARANG', 'Recency', 'Frequency', 'Monetary']
    return rfm


def categorize_rfm(rfm):
    # Menghitung quintile (5 bagian) untuk Recency, Frequency, dan Monetary
    recency_q1 = rfm['Recency'].quantile(0.2)
    recency_q2 = rfm['Recency'].quantile(0.4)
    recency_q3 = rfm['Recency'].quantile(0.6)
    recency_q4 = rfm['Recency'].quantile(0.8)
    
    frequency_q1 = rfm['Frequency'].quantile(0.2)
    frequency_q2 = rfm['Frequency'].quantile(0.4)
    frequency_q3 = rfm['Frequency'].quantile(0.6)
    frequency_q4 = rfm['Frequency'].quantile(0.8)
    
    monetary_q1 = rfm['Monetary'].quantile(0.2)
    monetary_q2 = rfm['Monetary'].quantile(0.4)
    monetary_q3 = rfm['Monetary'].quantile(0.6)
    monetary_q4 = rfm['Monetary'].quantile(0.8)

    # Membuat bins secara dinamis berdasarkan quintile
    recency_bins = [0, recency_q1, recency_q2, recency_q3, recency_q4, float('inf')]
    frequency_bins = [0, frequency_q1, frequency_q2, frequency_q3, frequency_q4, float('inf')]
    monetary_bins = [0, monetary_q1, monetary_q2, monetary_q3, monetary_q4, float('inf')]

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
    
    # Menampilkan kolom KODE BARANG dan NAMA BARANG, tanpa KATEGORI
    cluster_data = rfm[rfm['Cluster'] == cluster_label][['KODE BARANG', 'NAMA BARANG', 'Recency', 'Frequency', 'Monetary']]
    st.dataframe(cluster_data, width=400, height=350, key=f"cluster_table_{cluster_label}_{key_suffix}")

def process_category(rfm_category, category_name, n_clusters, key_suffix=''):
    # Memproses kategori dan menampilkan hasil
    if rfm_category.shape[0] > 0 and n_clusters > 0:
        scaler = StandardScaler()
        rfm_scaled = scaler.fit_transform(rfm_category[['Recency', 'Frequency', 'Monetary']])
        
        cluster_labels = cluster_rfm(rfm_scaled, n_clusters)
        rfm_category['Cluster'] = cluster_labels

        # Mengategorikan RFM tanpa argumen tambahan
        rfm_category = categorize_rfm(rfm_category)

        # Menghitung rata-rata RFM untuk setiap cluster
        cluster_means = rfm_category.groupby('Cluster')[['Recency', 'Frequency', 'Monetary']].mean()

        # Menghitung kuantile untuk menentukan kategori rata-rata RFM
        recency_quartiles = rfm_category['Recency'].quantile([0.2, 0.4, 0.6, 0.8])
        frequency_quartiles = rfm_category['Frequency'].quantile([0.2, 0.4, 0.6, 0.8])
        monetary_quartiles = rfm_category['Monetary'].quantile([0.2, 0.4, 0.6, 0.8])

        # Mendefinisikan fungsi untuk menentukan deskripsi berdasarkan kuantile yang disesuaikan
        def determine_category(value, quartiles, labels):
            if value <= quartiles[0.2]:
                return labels[0]
            elif value <= quartiles[0.4]:
                return labels[1]
            elif value <= quartiles[0.6]:
                return labels[2]
            elif value <= quartiles[0.8]:
                return labels[3]
            else:
                return labels[4]

        # Membuat legenda untuk setiap cluster berdasarkan rata-rata RFM
        custom_legends = {
            cluster: f"{determine_category(mean_values['Recency'], recency_quartiles, ['Baru Saja', 'Cukup Baru', 'Cukup Lama', 'Lama', 'Sangat Lama'])} Dibeli, "
                     f"Frekuensi {determine_category(mean_values['Frequency'], frequency_quartiles, ['Sangat Jarang', 'Jarang', 'Cukup Sering', 'Sering', 'Sangat Sering'])}, "
                     f"dan Nilai Pembelian {determine_category(mean_values['Monetary'], monetary_quartiles, ['Sangat Rendah', 'Rendah', 'Sedang', 'Tinggi', 'Sangat Tinggi'])}"
            for cluster, mean_values in cluster_means.iterrows()
        }

        col1, col2 = st.columns([1, 2])

        with col1:
            # Mengurangi ukuran font untuk judul kategori terjual
            st.markdown(f"<h4 style='font-size: 20px;'>Total {category_name} Terjual</h4>", unsafe_allow_html=True)
            st.markdown(f"<div style='border: 1px solid #d3d3d3; padding: 20px; border-radius: 10px;'>", unsafe_allow_html=True)

            # Mengurutkan clusters berdasarkan nama
            available_clusters = sorted(rfm_category['Cluster'].unique())

            for cluster_label in available_clusters:
                show_cluster_table(rfm_category, cluster_label, custom_legends.get(cluster_label, f'Cluster {cluster_label}'), key_suffix)

        with col2:
            st.markdown(f"<h4 style='font-size: 20px;'>Distribusi Cluster {category_name}</h4>", unsafe_allow_html=True)
            fig = plot_interactive_pie_chart(rfm_category, cluster_labels, category_name, custom_legends)
            st.plotly_chart(fig)

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