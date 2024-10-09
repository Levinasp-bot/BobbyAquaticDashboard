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
    # Menggunakan skala z-score untuk menghindari tumpang tindih pada kategori
    rfm['Recency_Z'] = (rfm['Recency'] - rfm['Recency'].mean()) / rfm['Recency'].std()
    rfm['Frequency_Z'] = (rfm['Frequency'] - rfm['Frequency'].mean()) / rfm['Frequency'].std()
    rfm['Monetary_Z'] = (rfm['Monetary'] - rfm['Monetary'].mean()) / rfm['Monetary'].std()

    # Menambahkan batas kategori z-score untuk 4 kategori
    bins = [-float('inf'), -0.5, 0.5, float('inf')]  # Menggunakan 4 kategori

    # Label untuk kategori
    rfm['Recency_Category'] = pd.cut(rfm['Recency_Z'], bins=bins, labels=['Sangat Baru', 'Cukup Lama', 'Sangat Lama'])
    rfm['Frequency_Category'] = pd.cut(rfm['Frequency_Z'], bins=bins, labels=['Jarang', 'Cukup Sering', 'Sering'])
    rfm['Monetary_Category'] = pd.cut(rfm['Monetary_Z'], bins=bins, labels=['Rendah', 'Sedang', 'Tinggi'])

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
    st.markdown(f"### Cluster: {custom_label}", unsafe_allow_html=True)
    
    cluster_data = rfm[rfm['Cluster'] == cluster_label]
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

        # Membuat legenda untuk setiap cluster menggunakan kategori linguistik
        custom_legends = {
            cluster: f"Recency: {rfm_category[rfm_category['Cluster'] == cluster]['Recency_Category'].mode()[0]}, "
                     f"Frequency: {rfm_category[rfm_category['Cluster'] == cluster]['Frequency_Category'].mode()[0]}, "
                     f"Monetary: {rfm_category[rfm_category['Cluster'] == cluster]['Monetary_Category'].mode()[0]}"
            for cluster in sorted(rfm_category['Cluster'].unique())
        }

        # Menampilkan informasi dalam dua kolom
        col1, col2 = st.columns(2)
        with col1:
            st.markdown(f"### Total {category_name} Terjual")
            st.markdown(f"<div style='border: 1px solid #d3d3d3; padding: 10px; border-radius: 5px;'>"
                        f"<strong>{rfm_category['Frequency'].sum()}</strong></div>", unsafe_allow_html=True)

        with col2:
            st.markdown("### Rata - rata RFM")
            average_rfm = rfm_category[['Recency', 'Frequency', 'Monetary']].mean()
            st.markdown(f"<div style='border: 1px solid #d3d3d3; padding: 10px; border-radius: 5px;'>"
                        f"<strong>Recency: {average_rfm['Recency']:.2f}</strong><br>"
                        f"<strong>Frequency: {average_rfm['Frequency']:.2f}</strong><br>"
                        f"<strong>Monetary: {average_rfm['Monetary']:.2f}</strong></div>", unsafe_allow_html=True)

        # Memilih cluster yang akan ditampilkan
        unique_key = f'selectbox_{category_name}_{key_suffix}_{str(hash(tuple(custom_legends.keys())))}'
        selected_custom_label = st.selectbox(
            f'Select a cluster for {category_name}:',
            options=[custom_legends[cluster] for cluster in sorted(custom_legends.keys())],
            key=unique_key
        )

        selected_cluster_num = {v: k for k, v in custom_legends.items()}[selected_custom_label]
        plot_key = f'plotly_chart_{category_name}_{key_suffix}'

        # Menampilkan grafik dan tabel cluster
        chart_col, table_col = st.columns(2)
        with chart_col:
            fig = plot_interactive_pie_chart(rfm_category, cluster_labels, category_name, custom_legends)
            st.plotly_chart(fig, use_container_width=True, key=plot_key)

        with table_col:
            show_cluster_table(rfm_category, selected_cluster_num, selected_custom_label, key_suffix=f'{category_name.lower()}_{selected_cluster_num}')

    else:
        st.error(f"Tidak ada data yang valid untuk clustering di kategori {category_name}.")

def get_optimal_k(data_scaled):
    # Mendapatkan jumlah cluster optimal menggunakan metode elbow
    model = KMeans(random_state=1)
    visualizer = KElbowVisualizer(model, k=(3, 10), timings=False)
    visualizer.fit(data_scaled)
    return visualizer.elbow_value_

def show_dashboard(data, key_suffix=''):
    # Menampilkan dashboard
    rfm = process_rfm(data)

    rfm_ikan = rfm[rfm['KATEGORI'] == 'Ikan']
    n_clusters_ikan = get_optimal_k(StandardScaler().fit_transform(rfm_ikan[['Recency', 'Frequency', 'Monetary']]))
    process_category(rfm_ikan, 'Ikan', n_clusters_ikan, key_suffix)

    rfm_aksesoris = rfm[rfm['KATEGORI'] == 'Aksesoris']
    n_clusters_aksesoris = get_optimal_k(StandardScaler().fit_transform(rfm_aksesoris[['Recency', 'Frequency', 'Monetary']]))
    process_category(rfm_aksesoris, 'Aksesoris', n_clusters_aksesoris, key_suffix)