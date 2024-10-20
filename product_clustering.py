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

def categorize_recency(recency, quartiles):
    if recency <= quartiles[0.2]:
        return 'Baru Saja'
    elif recency <= quartiles[0.4]:
        return 'Cukup Baru'
    elif recency <= quartiles[0.6]:
        return 'Cukup Lama'
    elif recency <= quartiles[0.8]:
        return 'Lama'
    else:
        return 'Sangat Lama'

def categorize_frequency(frequency, quartiles):
    if frequency <= quartiles[0.2]:
        return 'Sangat Jarang'
    elif frequency <= quartiles[0.4]:
        return 'Jarang'
    elif frequency <= quartiles[0.6]:
        return 'Cukup Sering'
    elif frequency <= quartiles[0.8]:
        return 'Sering'
    else:
        return 'Sangat Sering'

def categorize_monetary(monetary, quartiles):
    if monetary <= quartiles[0.2]:
        return 'Sangat Rendah'
    elif monetary <= quartiles[0.4]:
        return 'Rendah'
    elif monetary <= quartiles[0.6]:
        return 'Sedang'
    elif monetary <= quartiles[0.8]:
        return 'Tinggi'
    else:
        return 'Sangat Tinggi'

def categorize_rfm(rfm):
    # Menghitung quartiles untuk Recency, Frequency, dan Monetary
    recency_quartiles = rfm['Recency'].quantile([0.2, 0.4, 0.6, 0.8])
    frequency_quartiles = rfm['Frequency'].quantile([0.2, 0.4, 0.6, 0.8])
    monetary_quartiles = rfm['Monetary'].quantile([0.2, 0.4, 0.6, 0.8])

    # Kategorisasi Recency menggunakan fungsi categorize_recency
    rfm['Recency_Category'] = rfm['Recency'].apply(lambda x: categorize_recency(x, recency_quartiles))
    
    # Kategorisasi Frequency menggunakan fungsi categorize_frequency
    rfm['Frequency_Category'] = rfm['Frequency'].apply(lambda x: categorize_frequency(x, frequency_quartiles))
    
    # Kategorisasi Monetary menggunakan fungsi categorize_monetary
    rfm['Monetary_Category'] = rfm['Monetary'].apply(lambda x: categorize_monetary(x, monetary_quartiles))
    
    return rfm

def determine_cluster_characteristics(cluster_avg, recency_labels, frequency_labels, monetary_labels):
    def assign_cluster_characteristics(row):
        recency_cat = recency_labels[int(row['Recency'])]
        frequency_cat = frequency_labels[int(row['Frequency'])]
        monetary_cat = monetary_labels[int(row['Monetary'])]
        return pd.Series([recency_cat, frequency_cat, monetary_cat], index=['Recency_Category', 'Frequency_Category', 'Monetary_Category'])

    cluster_characteristics = cluster_avg.apply(assign_cluster_characteristics, axis=1)
    cluster_summary = pd.concat([cluster_avg, cluster_characteristics], axis=1)
    return cluster_summary

def cluster_rfm(rfm_scaled, n_clusters):
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
    st.markdown(f"##### Daftar Produk yang {custom_label}", unsafe_allow_html=True)
    
    cluster_data = rfm[rfm['Cluster'] == cluster_label]
    st.dataframe(cluster_data, width=400, height=350, key=f"cluster_table_{cluster_label}_{key_suffix}")

def process_category(rfm_category, category_name, n_clusters, key_suffix=''):
    if rfm_category.shape[0] > 0 and n_clusters > 0:
        scaler = StandardScaler()
        rfm_scaled = scaler.fit_transform(rfm_category[['Recency', 'Frequency', 'Monetary']])
        
        cluster_labels = cluster_rfm(rfm_scaled, n_clusters)
        rfm_category['Cluster'] = cluster_labels

        rfm_category = categorize_rfm(rfm_category)

        # Membuat summary rata-rata setiap cluster
        cluster_avg = rfm_category.groupby('Cluster')[['Recency', 'Frequency', 'Monetary']].mean()

        # Mendapatkan karakteristik cluster berdasarkan rata-rata RFM
        custom_legends = {
            cluster: f"{rfm_category[rfm_category['Cluster'] == cluster]['Recency'].mean():.2f} hari sejak pembelian terakhir, "
                    f"Rata-rata Frekuensi {rfm_category[rfm_category['Cluster'] == cluster]['Frequency'].mean():.2f}, "
                    f"dan Nilai Pembelian {rfm_category[rfm_category['Cluster'] == cluster]['Monetary'].mean():.2f}"
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
                        f"<span style='font-size: 32px; font-weight: bold;'>{average_rfm['Recency']:.2f}</span><br>"
                        f"<span style='font-size: 12px;'>Recency</span></div>"
                        f"<div style='text-align: center;'>"
                        f"<span style='font-size: 32px; font-weight: bold;'>{average_rfm['Frequency']:.2f}</span><br>"
                        f"<span style='font-size: 12px;'>Frequency</span></div>"
                        f"<div style='text-align: center;'>"
                        f"<span style='font-size: 32px; font-weight: bold;'>{average_rfm['Monetary']:.2f}</span><br>"
                        f"<span style='font-size: 12px;'>Monetary</span></div>"
                        f"</div>", unsafe_allow_html=True)

        unique_key = f'selectbox_{category_name}_{key_suffix}_{str(hash(tuple(custom_legends.keys())))}'
        selected_custom_label = st.selectbox(
            f'Pilih Kelompok untuk {category_name}:',
            options=[custom_legends[cluster] for cluster in sorted(custom_legends.keys())],
            key=unique_key
        )

        selected_cluster_num = {v: k for k, v in custom_legends.items()}[selected_custom_label]
        plot_key = f'plotly_chart_{category_name}_{key_suffix}'

        chart_col, table_col = st.columns(2)
        with chart_col:
            fig = plot_interactive_pie_chart(rfm_category, cluster_labels, category_name, custom_legends)
            st.plotly_chart(fig, use_container_width=True, key=plot_key)

        with table_col:
            show_cluster_table(rfm_category, selected_cluster_num, selected_custom_label, key_suffix=f'{category_name.lower()}_{selected_cluster_num}')

    else:
        st.error(f"Tidak ada data yang valid untuk clustering di kategori {category_name}.")

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
