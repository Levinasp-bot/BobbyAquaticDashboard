import glob
import os
import pandas as pd
from sklearn.preprocessing import StandardScaler
from sklearn.cluster import KMeans
import streamlit as st
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

def calculate_wcss(rfm_scaled, max_clusters=10):
    """Calculate WCSS for a range of cluster numbers to determine the optimal k."""
    wcss = []
    for i in range(1, max_clusters+1):
        kmeans = KMeans(n_clusters=i, init='k-means++', random_state=1)
        kmeans.fit(rfm_scaled)
        wcss.append(kmeans.inertia_)
    return wcss

def find_optimal_k(wcss):
    """Find the optimal number of clusters using the Elbow Method."""
    k_opt = 1
    # A simple heuristic: Find the point where WCSS drops sharply.
    # This logic can be refined for more complex data.
    for i in range(1, len(wcss)-1):
        if (wcss[i-1] - wcss[i]) > (wcss[i] - wcss[i+1]):
            k_opt = i + 1
            break
    return k_opt

def cluster_rfm(rfm_scaled):
    """Automatically determine the best k using the Elbow Method and cluster the data."""
    wcss = calculate_wcss(rfm_scaled)
    optimal_k = find_optimal_k(wcss)
    kmeans = KMeans(n_clusters=optimal_k, init='k-means++', random_state=1)
    kmeans.fit(rfm_scaled)
    return kmeans.labels_, optimal_k

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

def generate_custom_legends(n_clusters, category_name):
    """Template to generate cluster labels based on the category name."""
    template = f'{category_name} Kualitas {{}}'
    return {i: template.format(i + 1) for i in range(n_clusters)}

def show_cluster_table(rfm, cluster_label, custom_label, key_suffix):
    st.markdown(f"<h4>Cluster: {custom_label} Members</h4>", unsafe_allow_html=True)
    cluster_data = rfm[rfm['Cluster'] == cluster_label]
    st.dataframe(cluster_data, width=400, height=350, key=f"cluster_table_{cluster_label}_{key_suffix}")

def process_category(rfm_category, category_name, key_suffix=''):
    if rfm_category.shape[0] > 0:
        scaler = StandardScaler()
        rfm_scaled = scaler.fit_transform(rfm_category[['Recency', 'Frequency', 'Monetary']])
        
        # Automatically determine the number of clusters and assign labels
        cluster_labels, n_clusters = cluster_rfm(rfm_scaled)
        rfm_category['Cluster'] = cluster_labels

        # Generate custom legends based on the determined number of clusters
        custom_legends = generate_custom_legends(n_clusters, category_name)

        # Display total items sold and average RFM values
        total_fish_sold = rfm_category['Frequency'].sum() if category_name == 'Ikan' else 0
        total_accessories_sold = rfm_category['Frequency'].sum() if category_name == 'Aksesoris' else 0
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
                             f"<strong>{total_accessories_sold}</strong></div>")
        
        with col2:
            st.markdown("### Rata - rata RFM")
            st.markdown(f"<div style='border: 1px solid #d3d3d3; padding: 10px; border-radius: 5px;'>"
                         f"<strong>Recency: {average_rfm['Recency']:.2f}</strong><br>"
                         f"<strong>Frequency: {average_rfm['Frequency']:.2f}</strong><br>"
                         f"<strong>Monetary: {average_rfm['Monetary']:.2f}</strong></div>", unsafe_allow_html=True)

        unique_key = f'selectbox_{category_name}_{key_suffix}_{str(hash(tuple(cluster_labels)))}'

        selected_custom_label = st.selectbox(
            f'Select a cluster for {category_name}:',
            options=[custom_legends[cluster] for cluster in sorted(set(cluster_labels))],
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

def show_dashboard(data, key_suffix=''):
    rfm = process_rfm(data)

    rfm_ikan = rfm[rfm['KATEGORI'] == 'Ikan']
    rfm_aksesoris = rfm[rfm['KATEGORI'] == 'Aksesoris']

    # Process each category dynamically with flexible k (determined from the data)
    process_category(rfm_ikan, 'Ikan', key_suffix='ikan')
    process_category(rfm_aksesoris, 'Aksesoris', key_suffix='aksesoris')
