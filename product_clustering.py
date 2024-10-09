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
    # Mengkategorikan RFM berdasarkan kuartil
    for category in rfm['KATEGORI'].unique():
        category_data = rfm[rfm['KATEGORI'] == category]

        if not category_data.empty:
            Q1_recency = category_data['Recency'].quantile(0.25)
            Q2_recency = category_data['Recency'].quantile(0.5)  # Median
            Q3_recency = category_data['Recency'].quantile(0.75)

            Q1_frequency = category_data['Frequency'].quantile(0.25)
            Q2_frequency = category_data['Frequency'].quantile(0.5)
            Q3_frequency = category_data['Frequency'].quantile(0.75)

            Q1_monetary = category_data['Monetary'].quantile(0.25)
            Q2_monetary = category_data['Monetary'].quantile(0.5)
            Q3_monetary = category_data['Monetary'].quantile(0.75)

            # Menentukan kategori berdasarkan kuartil
            rfm.loc[(rfm['KATEGORI'] == category) & (rfm['Recency'] <= Q1_recency), 'Recency_Category'] = 'Baru Saja'
            rfm.loc[(rfm['KATEGORI'] == category) & (rfm['Recency'].between(Q1_recency, Q2_recency)), 'Recency_Category'] = 'Cukup Lama'
            rfm.loc[(rfm['KATEGORI'] == category) & (rfm['Recency'] > Q2_recency), 'Recency_Category'] = 'Sangat Lama'

            rfm.loc[(rfm['KATEGORI'] == category) & (rfm['Frequency'] <= Q1_frequency), 'Frequency_Category'] = 'Jarang'
            rfm.loc[(rfm['KATEGORI'] == category) & (rfm['Frequency'].between(Q1_frequency, Q2_frequency)), 'Frequency_Category'] = 'Cukup Sering'
            rfm.loc[(rfm['KATEGORI'] == category) & (rfm['Frequency'] > Q2_frequency), 'Frequency_Category'] = 'Sering'

            rfm.loc[(rfm['KATEGORI'] == category) & (rfm['Monetary'] <= Q1_monetary), 'Monetary_Category'] = 'Rendah'
            rfm.loc[(rfm['KATEGORI'] == category) & (rfm['Monetary'].between(Q1_monetary, Q2_monetary)), 'Monetary_Category'] = 'Sedang'
            rfm.loc[(rfm['KATEGORI'] == category) & (rfm['Monetary'] > Q2_monetary), 'Monetary_Category'] = 'Tinggi'

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

        rfm_category = categorize_rfm(rfm_category)

        # Membuat legenda untuk setiap cluster
        custom_legends = {
            cluster: f"Recency {rfm_category[rfm_category['Cluster'] == cluster]['Recency_Category'].mode()[0]}, "
                     f"Frequency {rfm_category[rfm_category['Cluster'] == cluster]['Frequency_Category'].mode()[0]}, "
                     f" dan Monetary {rfm_category[rfm_category['Cluster'] == cluster]['Monetary_Category'].mode()[0]}"
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
                    # Menampilkan grafik pie interaktif
            fig = plot_interactive_pie_chart(rfm_category, rfm_category['Cluster'], category_name, custom_legends)
            st.plotly_chart(fig, use_container_width=True, key=plot_key)

        with table_col:
            # Menampilkan tabel cluster
            show_cluster_table(rfm_category, selected_cluster_num, selected_custom_label, key_suffix)
