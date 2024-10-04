import glob
import os
import pandas as pd
from sklearn.preprocessing import StandardScaler
from sklearn.cluster import KMeans
import streamlit as st
import plotly.graph_objects as go

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

    # Adjusting pie chart size and legend for better display
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
        ),
        height=400,  # Adjust height of the figure
        width=400    # Adjust width of the figure
    )

    return fig

def show_cluster_table(rfm, cluster_label, custom_label, key_suffix):
    # Use markdown with HTML to adjust the title size instead of st.subheader
    st.markdown(f"<h4>Cluster: {custom_label} Members</h4>", unsafe_allow_html=True)
    
    cluster_data = rfm[rfm['Cluster'] == cluster_label]
    
    # Adjust the width and height of the dataframe to fit better
    st.dataframe(cluster_data, width=400, height=350, key=f"cluster_table_{cluster_label}_{key_suffix}")

def process_category(rfm_category, category_name, n_clusters, custom_legends, key_suffix=''):
    if rfm_category.shape[0] > 0:
        scaler = StandardScaler()
        rfm_scaled = scaler.fit_transform(rfm_category[['Recency', 'Frequency', 'Monetary']])
        
        cluster_labels = cluster_rfm(rfm_scaled, n_clusters)
        rfm_category['Cluster'] = cluster_labels

        available_clusters = sorted(rfm_category['Cluster'].unique())
        custom_label_map = {cluster: custom_legends.get(cluster, f'Cluster {cluster}') for cluster in available_clusters}

        # Calculate Total Products Sold and Average RFM
        total_products_sold = rfm_category['Frequency'].sum()
        average_rfm = rfm_category[['Recency', 'Frequency', 'Monetary']].mean()

        # Adjust layout for side-by-side display
        col1, col2 = st.columns(2)  # Create two equal columns

        # Display total products sold and average RFM
        with col1:
            st.markdown("### Total Products Sold")
            st.markdown(f"<div style='border: 1px solid #d3d3d3; padding: 10px; border-radius: 5px;'>"
                         f"<strong>{total_products_sold}</strong></div>", unsafe_allow_html=True)
        
        with col2:
            st.markdown("### Average RFM")
            st.markdown(f"<div style='border: 1px solid #d3d3d3; padding: 10px; border-radius: 5px;'>"
                         f"<strong>Recency: {average_rfm['Recency']:.2f}</strong><br>"
                         f"<strong>Frequency: {average_rfm['Frequency']:.2f}</strong><br>"
                         f"<strong>Monetary: {average_rfm['Monetary']:.2f}</strong></div>", unsafe_allow_html=True)

        # Create a unique key for the selectbox
        unique_key = f'selectbox_{category_name}_{key_suffix}_{str(hash(tuple(available_clusters)))}'

        selected_custom_label = st.selectbox(
            f'Select a cluster for {category_name}:',
            options=[custom_label_map[cluster] for cluster in available_clusters],
            key=unique_key
        )

        selected_cluster_num = {v: k for k, v in custom_label_map.items()}[selected_custom_label]

        # Create a unique key for the plotly chart
        plot_key = f'plotly_chart_{category_name}_{key_suffix}'

        # Create columns for chart and table
        chart_col, table_col = st.columns(2)  # Create two equal columns for chart and table

        with chart_col:
            fig = plot_interactive_pie_chart(rfm_category, cluster_labels, category_name, custom_legends)
            st.plotly_chart(fig, use_container_width=True, key=plot_key)  # Pie chart in the first column

        with table_col:
            show_cluster_table(rfm_category, selected_cluster_num, selected_custom_label, key_suffix=f'{category_name.lower()}_{selected_cluster_num}')

    else:
        st.error(f"Tidak ada data yang valid untuk clustering di kategori {category_name}.")


def show_dashboard(data, key_suffix=''):
    rfm = process_rfm(data)

    rfm_ikan = rfm[rfm['KATEGORI'] == 'Ikan']
    rfm_aksesoris = rfm[rfm['KATEGORI'] == 'Aksesoris']

    k_ikan = 4
    k_aksesoris = 4

    custom_legends = {
        'Ikan': {0: 'Ikan Kualitas Tinggi', 1: 'Ikan Kualitas Menengah', 2: 'Ikan Kualitas Rendah', 3: 'Ikan Spesial'},
        'Aksesoris': {0: 'Aksesoris Populer', 1: 'Aksesoris Baru', 2: 'Aksesoris Diskon', 3: 'Aksesoris Premium'}
    }

    process_category(rfm_ikan, 'Ikan', k_ikan, custom_legends['Ikan'], key_suffix='ikan_file2')  # Kunci unik untuk file 2
    process_category(rfm_aksesoris, 'Aksesoris', k_aksesoris, custom_legends['Aksesoris'], key_suffix='aksesoris_file2')