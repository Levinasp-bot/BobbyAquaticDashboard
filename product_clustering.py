import glob
import os
import pandas as pd
from sklearn.preprocessing import StandardScaler
from sklearn.cluster import KMeans
import streamlit as st
import plotly.express as px

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
    
    rfm = data.groupby('KODE BARANG').agg({
        'TANGGAL': lambda x: (reference_date - x.max()).days,  # Recency
        'NAMA BARANG': 'count',  
        'TOTAL HR JUAL': 'sum'  
    }).reset_index()
    
    rfm.columns = ['KODE BARANG', 'Recency', 'Frequency', 'Monetary']
    return rfm

@st.cache
def cluster_rfm(rfm_scaled, n_clusters):
    kmeans = KMeans(n_clusters=n_clusters, init='k-means++', random_state=1)
    kmeans.fit(rfm_scaled)
    return kmeans.labels_

def plot_interactive_pie_chart(rfm, cluster_labels):
    rfm['Cluster'] = cluster_labels
    cluster_counts = rfm['Cluster'].value_counts().reset_index()
    cluster_counts.columns = ['Cluster', 'Count']

    fig = px.pie(cluster_counts, values='Count', names='Cluster', title='Cluster Distribution', hole=0.3)
    fig.update_traces(textinfo='percent+label', pull=[0.05]*len(cluster_counts))

    st.plotly_chart(fig, use_container_width=True)
    return rfm

def show_cluster_table(rfm, cluster_label):
    st.subheader(f"Cluster {cluster_label} Members")
    cluster_data = rfm[rfm['Cluster'] == cluster_label]
    st.dataframe(cluster_data)

def show_dashboard(data, n_clusters, key_suffix=''):
    rfm = process_rfm(data)

    scaler = StandardScaler()
    rfm_scaled = scaler.fit_transform(rfm[['Recency', 'Frequency', 'Monetary']])

    if rfm_scaled.shape[0] > 0:
        cluster_labels = cluster_rfm(rfm_scaled, n_clusters)

        st.subheader(f"Cluster Distribution Visualization{key_suffix}")
        rfm_with_clusters = plot_interactive_pie_chart(rfm, cluster_labels)

        # Interactivity: Select a cluster from the pie chart
        cluster_to_show = st.selectbox('Select a cluster to view its members:', sorted(rfm_with_clusters['Cluster'].unique()))
        show_cluster_table(rfm_with_clusters, cluster_to_show)
    else:
        st.error("Tidak ada data yang valid untuk clustering.")
