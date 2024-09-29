import pandas as pd
import numpy as np
import plotly.express as px
import streamlit as st
from sklearn.preprocessing import StandardScaler
from sklearn.cluster import KMeans

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

def plot_pie_chart(cluster_labels, rfm):
    cluster_counts = pd.Series(cluster_labels).value_counts()
    
    # Create a pie chart using Plotly
    fig = px.pie(cluster_counts, values=cluster_counts.values, names=cluster_counts.index,
                 title="Cluster Distribution", hole=0.3)
    
    st.plotly_chart(fig)

    # Return the cluster labels with the rfm data for later use
    rfm['Cluster'] = cluster_labels
    return rfm

def display_cluster_members(rfm, cluster_selected):
    st.subheader(f"Cluster {cluster_selected} Members")
    cluster_data = rfm[rfm['Cluster'] == cluster_selected]
    st.write(cluster_data)

def show_dashboard(data, key_suffix=''):
    st.subheader("Data Overview")
    st.write(data)

    rfm = process_rfm(data)

    scaler = StandardScaler()
    rfm_scaled = scaler.fit_transform(rfm[['Recency', 'Frequency', 'Monetary']])

    if rfm_scaled.shape[0] > 0:
        # Define the number of clusters (can use KMeans Elbow method here)
        n_clusters = 3  # Example fixed number of clusters
        cluster_labels = cluster_rfm(rfm_scaled, n_clusters)

        st.subheader(f"Cluster Distribution Visualization{key_suffix}")
        rfm_with_clusters = plot_pie_chart(cluster_labels, rfm)

        # Let the user click on the pie chart and display corresponding cluster members
        cluster_selected = st.selectbox("Select Cluster to view members", sorted(rfm_with_clusters['Cluster'].unique()))

        # Show the table of members in the selected cluster
        display_cluster_members(rfm_with_clusters, cluster_selected)
    else:
        st.error("Tidak ada data yang valid untuk clustering.")
