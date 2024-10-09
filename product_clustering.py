import glob
import os
import pandas as pd
from sklearn.preprocessing import StandardScaler
from sklearn.cluster import KMeans
import streamlit as st
import plotly.graph_objects as go
from yellowbrick.cluster import KElbowVisualizer

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

def process_rfm(data):
    data['TANGGAL'] = pd.to_datetime(data['TANGGAL'])
    reference_date = data['TANGGAL'].max()
    
    # Group by KODE BARANG and KATEGORI
    rfm = data.groupby(['KODE BARANG', 'KATEGORI']).agg({
        'TANGGAL': lambda x: (reference_date - x.max()).days,  # Recency
        'NAMA BARANG': 'count',  # Frequency
        'TOTAL HR JUAL': 'sum'  # Monetary
    }).reset_index()
    
    rfm.columns = ['KODE BARANG', 'KATEGORI', 'Recency', 'Frequency', 'Monetary']
    return rfm

def categorize_rfm(rfm, category_name):
    if category_name == 'Ikan':
        recency_bins = [0, 3, 16.75, float('inf')]
        frequency_bins = [0, 82, 146.5, float('inf')]
        monetary_bins = [0, 3540625, 11900000, float('inf')]
        recency_labels = ['Baru Saja', 'Cukup Lama', 'Sangat Lama']
        frequency_labels = ['Jarang', 'Cukup Sering', 'Sering']
        monetary_labels = ['Rendah', 'Sedang', 'Tinggi']
    elif category_name == 'Aksesoris':
        recency_bins = [0, 16, 182.5, float('inf')]
        frequency_bins = [0, 8, 60.25, float('inf')]
        monetary_bins = [0, 1007500, 3456250, float('inf')]
        recency_labels = ['Baru Saja', 'Cukup Lama', 'Sangat Lama']
        frequency_labels = ['Jarang', 'Cukup Sering', 'Sering']
        monetary_labels = ['Rendah', 'Sedang', 'Tinggi']
    else:
        return rfm

    rfm['Recency_Category'] = pd.cut(rfm['Recency'], bins=recency_bins, labels=recency_labels)
    rfm['Frequency_Category'] = pd.cut(rfm['Frequency'], bins=frequency_bins, labels=frequency_labels)
    rfm['Monetary_Category'] = pd.cut(rfm['Monetary'], bins=monetary_bins, labels=monetary_labels)
    
    return rfm

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

def show_cluster_table(rfm, cluster_label, custom_label, key_suffix):
    st.markdown(f"<h4>Cluster dengan {custom_label} </h4>", unsafe_allow_html=True)
    
    cluster_data = rfm[rfm['Cluster'] == cluster_label]
    st.dataframe(cluster_data, width=400, height=350, key=f"cluster_table_{cluster_label}_{key_suffix}")

def process_category(rfm_category, category_name, n_clusters, key_suffix=''):
    if rfm_category.shape[0] > 0:
        scaler = StandardScaler()
        rfm_scaled = scaler.fit_transform(rfm_category[['Recency', 'Frequency', 'Monetary']])
        
        cluster_labels = cluster_rfm(rfm_scaled, n_clusters)
        rfm_category['Cluster'] = cluster_labels

        rfm_category = categorize_rfm(rfm_category, category_name)

        custom_legends = {
            cluster: f"Recency {rfm_category[rfm_category['Cluster'] == cluster]['Recency_Category'].mode()[0]}, "
                     f"Frequency {rfm_category[rfm_category['Cluster'] == cluster]['Frequency_Category'].mode()[0]}, "
                     f" dan Monetary {rfm_category[rfm_category['Cluster'] == cluster]['Monetary_Category'].mode()[0]}"
            for cluster in sorted(rfm_category['Cluster'].unique())
        }

        if category_name == 'Ikan':
            total_fish_sold = rfm_category['Frequency'].sum()
            total_accessories_sold = 0
        elif category_name == 'Aksesoris':
            total_accessories_sold = rfm_category['Frequency'].sum()
            total_fish_sold = 0

        average_rfm = rfm_category[['Recency', 'Frequency', 'Monetary']].mean()

        col1, col2 = st.columns(2)
        with col1:
            if category_name == 'Ikan':
                st.markdown("### Total Ikan Terjual")
                st.markdown(f"<div style='border: 1px solid #d3d3d3; padding: 10px; border-radius: 5px;'>"
                            f"<strong>{total_fish_sold}</strong></div>", unsafe_allow_html=True)
            else:
                st.markdown("### Total Aksesoris Terjual")
                st.markdown(f"<div style='border: 1px solid #d3d3d3; padding: 10px; border-radius: 5px;'>"
                            f"<strong>{total_accessories_sold}</strong></div>", unsafe_allow_html=True)
        
        with col2:
            st.markdown("### Rata - rata RFM")
            st.markdown(f"<div style='border: 1px solid #d3d3d3; padding: 10px; border-radius: 5px;'>"
                        f"<strong>Recency: {average_rfm['Recency']:.2f}</strong><br>"
                        f"<strong>Frequency: {average_rfm['Frequency']:.2f}</strong><br>"
                        f"<strong>Monetary: {average_rfm['Monetary']:.2f}</strong></div>", unsafe_allow_html=True)

        unique_key = f'selectbox_{category_name}_{key_suffix}_{str(hash(tuple(custom_legends.keys())))}'
        selected_custom_label = st.selectbox(
            f'Select a cluster for {category_name}:',
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
    visualizer = KElbowVisualizer(model, k=(3, 10), timings=False)
    visualizer.fit(data_scaled)
    return visualizer.elbow_value_

def show_dashboard(data, key_suffix=''):
    rfm = process_rfm(data)

    rfm_ikan = rfm[rfm['KATEGORI'] == 'Ikan']
    rfm_aksesoris = rfm[rfm['KATEGORI'] == 'Aksesoris']

    n_clusters_ikan = get_optimal_k(StandardScaler().fit_transform(rfm_ikan[['Recency', 'Frequency', 'Monetary']])) if not rfm_ikan.empty else 0
    n_clusters_aksesoris = get_optimal_k(StandardScaler().fit_transform(rfm_aksesoris[['Recency', 'Frequency', 'Monetary']])) if not rfm_aksesoris.empty else 0

    if n_clusters_ikan > 0:
        process_category(rfm_ikan, 'Ikan', n_clusters_ikan, key_suffix=key_suffix)
    if n_clusters_aksesoris > 0:
        process_category(rfm_aksesoris, 'Aksesoris', n_clusters_aksesoris, key_suffix=key_suffix)
