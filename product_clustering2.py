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

@st.cache
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

@st.cache
def cluster_rfm(rfm_scaled, n_clusters):
    kmeans = KMeans(n_clusters=n_clusters, init='k-means++', random_state=1)
    kmeans.fit(rfm_scaled)
    return kmeans.labels_

def plot_interactive_pie_chart(rfm, cluster_labels, category_name, custom_legends):
    rfm['Cluster'] = cluster_labels
    cluster_counts = rfm['Cluster'].value_counts().reset_index()
    cluster_counts.columns = ['Cluster', 'Count']

    # Ensure custom legends map to actual clusters generated
    available_clusters = sorted(rfm['Cluster'].unique())
    custom_legend_mapped = {cluster: custom_legends.get(cluster, f'Cluster {cluster}') for cluster in available_clusters}

    # Apply the custom legends mapping
    cluster_counts['Cluster'] = cluster_counts['Cluster'].map(custom_legend_mapped)

    # Create the pie chart with custom labels
    fig = go.Figure(data=[go.Pie(
        labels=cluster_counts['Cluster'],  # Use custom labels here
        values=cluster_counts['Count'],
        hole=0.3,
        textinfo='percent+label',
        pull=[0.05] * len(cluster_counts)
    )])

    # Customize the legend
    fig.update_layout(
        legend=dict(
            itemsizing='constant',
            orientation='h',  # Horizontal
            yanchor='bottom',
            y=1.02,  # Position legend at the top
            xanchor='center',
            x=0.5,
            traceorder='normal'
        )
    )

    return fig

def show_cluster_table(rfm, cluster_label, custom_label, key_suffix):
    st.subheader(f"Cluster: {custom_label} Members")
    cluster_data = rfm[rfm['Cluster'] == cluster_label]
    st.dataframe(cluster_data, key=f"cluster_table_{cluster_label}_{key_suffix}")

def process_category(rfm_category, category_name, n_clusters, custom_legends, key_suffix=''):
    if rfm_category.shape[0] > 0:
        # Scale the RFM data
        scaler = StandardScaler()
        rfm_scaled = scaler.fit_transform(rfm_category[['Recency', 'Frequency', 'Monetary']])
        
        # Cluster the data with defined k
        cluster_labels = cluster_rfm(rfm_scaled, n_clusters)
        
        # Add cluster labels to the dataframe
        rfm_category['Cluster'] = cluster_labels

        # Custom label mapping for selectbox and table
        available_clusters = sorted(rfm_category['Cluster'].unique())
        custom_label_map = {cluster: custom_legends.get(cluster, f'Cluster {cluster}') for cluster in available_clusters}
        
        # Create two columns with 2/3 for chart and 1/3 for the table
        col1, col2 = st.columns([2, 1])  # Resize: 2 for chart, 1 for table

        # Selectbox with custom label for selecting cluster
        selected_custom_label = col1.selectbox(
            f'Select a cluster for {category_name}:', 
            options=[custom_label_map[cluster] for cluster in available_clusters], 
            key=f'selectbox_{category_name}_{key_suffix}'
        )

        # Find the numeric cluster corresponding to the custom label
        selected_cluster_num = {v: k for k, v in custom_label_map.items()}[selected_custom_label]

        # Pie chart in the left column
        fig = plot_interactive_pie_chart(rfm_category, cluster_labels, category_name, custom_legends)
        col1.plotly_chart(fig, use_container_width=True)

        # Detailed table in the right column with custom labels
        show_cluster_table(rfm_category, selected_cluster_num, selected_custom_label, key_suffix=f'{category_name.lower()}_{selected_cluster_num}')
    else:
        st.error(f"Tidak ada data yang valid untuk clustering di kategori {category_name}.")

def show_dashboard(data, key_suffix=''):
    # Process RFM for the entire dataset
    rfm = process_rfm(data)

    # Filter RFM data by category
    rfm_ikan = rfm[rfm['KATEGORI'] == 'Ikan']
    rfm_aksesoris = rfm[rfm['KATEGORI'] == 'Aksesoris']

    # Define the number of clusters manually for each category
    k_ikan = 4  # Number of clusters for 'Ikan'
    k_aksesoris = 4  # Number of clusters for 'Aksesoris'

    # Custom legends
    custom_legends = {
        'Ikan': {0: 'Ikan Kualitas Tinggi', 1: 'Ikan Kualitas Menengah', 2: 'Ikan Kualitas Rendah', 3: 'Ikan Spesial'},
        'Aksesoris': {0: 'Aksesoris Populer', 1: 'Aksesoris Baru', 2: 'Aksesoris Diskon', 3: 'Aksesoris Premium'}
    }

    # Process clustering for 'Ikan' and 'Aksesoris' with predefined k values
    process_category(rfm_ikan, 'Ikan', k_ikan, custom_legends['Ikan'], key_suffix='ikan')
    process_category(rfm_aksesoris, 'Aksesoris', k_aksesoris, custom_legends['Aksesoris'], key_suffix='aksesoris')
