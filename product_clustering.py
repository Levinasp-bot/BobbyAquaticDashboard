import glob
import os
import pandas as pd
from sklearn.preprocessing import StandardScaler
from sklearn.cluster import KMeans
import streamlit as st
import plotly.graph_objects as go
from yellowbrick.cluster import KElbowVisualizer

@st.cache_data
def load_all_excel_files(folder_path, sheet_name):
    all_files = glob.glob(os.path.join(folder_path, "*.xlsm"))
    dfs = []
    for file in all_files:
        df = pd.read_excel(file, sheet_name=sheet_name)
        if 'KODE BARANG' in df.columns:
            df = df.loc[:, ~df.columns.duplicated()]
        dfs.append(df)
    return pd.concat(dfs, ignore_index=True)

def process_rfm(data):
    data['TANGGAL'] = pd.to_datetime(data['TANGGAL'])
    reference_date = data['TANGGAL'].max()
    rfm = data.groupby(['KODE BARANG', 'KATEGORI']).agg({
        'TANGGAL': lambda x: (reference_date - x.max()).days,  # Recency
        'NAMA BARANG': 'count',  # Frequency
        'TOTAL HR JUAL': 'sum'  # Monetary
    }).reset_index()
    rfm.columns = ['KODE BARANG', 'KATEGORI', 'Recency', 'Frequency', 'Monetary']
    return rfm

def categorize_rfm(rfm):
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

    recency_bins = [0, recency_q1, recency_q2, recency_q3, recency_q4, float('inf')]
    frequency_bins = [0, frequency_q1, frequency_q2, frequency_q3, frequency_q4, float('inf')]
    monetary_bins = [0, monetary_q1, monetary_q2, monetary_q3, monetary_q4, float('inf')]

    recency_labels = ['Baru Saja', 'Cukup Baru', 'Cukup Lama', 'Lama', 'Sangat Lama']
    frequency_labels = ['Sangat Jarang', 'Jarang', 'Cukup Sering', 'Sering', 'Sangat Sering']
    monetary_labels = ['Sangat Rendah', 'Rendah', 'Sedang', 'Tinggi', 'Sangat Tinggi']

    rfm['Recency_Category'] = pd.cut(rfm['Recency'], bins=recency_bins, labels=recency_labels)
    rfm['Frequency_Category'] = pd.cut(rfm['Frequency'], bins=frequency_bins, labels=frequency_labels)
    rfm['Monetary_Category'] = pd.cut(rfm['Monetary'], bins=monetary_bins, labels=monetary_labels)

    return rfm

def assign_cluster_labels(cluster_avg, recency_labels, frequency_labels, monetary_labels):
    def assign_linguistic_labels(value, labels):
        for i, label in enumerate(labels):
            if value <= i + 1:
                return label
        return labels[-1]

    cluster_labels = {}
    for cluster, row in cluster_avg.iterrows():
        recency_cat = assign_linguistic_labels(row['Recency'], recency_labels)
        frequency_cat = assign_linguistic_labels(row['Frequency'], frequency_labels)
        monetary_cat = assign_linguistic_labels(row['Monetary'], monetary_labels)

        cluster_labels[cluster] = f"{recency_cat} Dibeli, Frekuensi {frequency_cat}, dan Nilai Pembelian {monetary_cat}"

    return cluster_labels

def cluster_rfm(rfm_scaled, n_clusters):
    kmeans = KMeans(n_clusters=n_clusters, init='k-means++', random_state=1)
    kmeans.fit(rfm_scaled)
    return kmeans.labels_

def plot_interactive_pie_chart(rfm, cluster_labels, category_name, custom_legends):
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

def show_cluster_table(rfm, cluster_label, custom_label, key_suffix, table_index):
    st.markdown(f"##### Daftar Produk yang {custom_label}", unsafe_allow_html=True)
    
    cluster_data = rfm[rfm['Cluster'] == cluster_label]
    st.dataframe(cluster_data, width=400, height=350, key=f"cluster_table_{cluster_label}_{key_suffix}_{table_index}")

def process_category(rfm_category, category_name, n_clusters, key_suffix=''):
    if rfm_category.shape[0] > 0 and n_clusters > 0:
        scaler = StandardScaler()
        rfm_scaled = scaler.fit_transform(rfm_category[['Recency', 'Frequency', 'Monetary']])
        
        cluster_labels = cluster_rfm(rfm_scaled, n_clusters)
        rfm_category['Cluster'] = cluster_labels

        rfm_category = categorize_rfm(rfm_category)

        # Membuat summary rata-rata setiap cluster
        cluster_avg = rfm_category.groupby('Cluster')[['Recency', 'Frequency', 'Monetary']].mean()

        # Menentukan label berdasarkan rata-rata RFM
        custom_legends = assign_cluster_labels(cluster_avg, ['Baru Saja', 'Cukup Baru', 'Cukup Lama', 'Lama', 'Sangat Lama'],
                                               ['Sangat Jarang', 'Jarang', 'Cukup Sering', 'Sering', 'Sangat Sering'],
                                               ['Sangat Rendah', 'Rendah', 'Sedang', 'Tinggi', 'Sangat Tinggi'])

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
                        f"<span style='font-size: 32px; font-weight: bold;'>{average_rfm['Recency']:.2f}</span><br>"
                        f"<span style='font-size: 12px;'>Recency</span></div>"
                        f"<div style='text-align: center;'>"
                        f"<span style='font-size: 32px; font-weight: bold;'>{average_rfm['Frequency']:.2f}</span><br>"
                        f"<span style='font-size: 12px;'>Frequency</span></div>"
                        f"<div style='text-align: center;'>"
                        f"<span style='font-size: 32px; font-weight: bold;'>{average_rfm['Monetary']:.2f}</span><br>"
                        f"<span style='font-size: 12px;'>Monetary</span></div>"
                        f"</div>", unsafe_allow_html=True)

        pie_chart = plot_interactive_pie_chart(rfm_category, cluster_labels, category_name, custom_legends)
        st.plotly_chart(pie_chart, use_container_width=True)

        if st.checkbox(f"Lihat Detail Cluster {category_name}", key=f"cluster_detail_{key_suffix}"):
            cluster_selected = st.selectbox(f"Pilih Cluster {category_name}", sorted(custom_legends.keys()),
                                            format_func=lambda x: custom_legends[x], key=f"select_cluster_{key_suffix}")

            # Menggunakan index untuk membedakan tabel
            show_cluster_table(rfm_category, cluster_selected, custom_legends[cluster_selected], key_suffix, table_index=1)  # Tabel pertama
            show_cluster_table(rfm_category, cluster_selected, custom_legends[cluster_selected], key_suffix, table_index=2)  # Tabel kedua

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
