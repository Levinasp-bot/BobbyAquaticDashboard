import glob
import os
import pandas as pd
from sklearn.preprocessing import StandardScaler
from sklearn.cluster import KMeans
import matplotlib.pyplot as plt
import seaborn as sns
import streamlit as st
from yellowbrick.cluster import KElbowVisualizer

@st.cache
def load_all_excel_files(folder_path, sheet_name):
    all_files = glob.glob(os.path.join(folder_path, "*.xlsm"))
    dfs = []
    for file in all_files:
        df = pd.read_excel(file, sheet_name=sheet_name)
        if 'KODE BARANG' in df.columns:
            df = df.loc[:, ~df.columns.duplicated()]  # Remove duplicate columns
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
def find_optimal_k(rfm_scaled):
    if rfm_scaled.shape[0] <= 1:
        st.error("Data tidak cukup untuk menemukan jumlah cluster. Pastikan data tidak kosong.")
        return None  

    model = KMeans(random_state=1)
    visualizer = KElbowVisualizer(model, k=(2, 10), timings=True)
    
    try:
        visualizer.fit(rfm_scaled)
        visualizer.show()
        optimal_k = visualizer.elbow_value_  # Get the optimal K from elbow method
        return optimal_k
    except Exception as e:
        st.error(f"Error during Elbow Method visualization: {str(e)}")
        return None

@st.cache
def cluster_rfm(rfm_scaled, n_clusters):
    kmeans = KMeans(n_clusters=n_clusters, init='k-means++', random_state=1)
    kmeans.fit(rfm_scaled)
    return kmeans.labels_

def plot_3d_clusters(rfm_scaled, cluster_labels):
    fig = plt.figure(figsize=(7, 7))
    ax = fig.add_subplot(111, projection='3d')
    
    scatter = ax.scatter(rfm_scaled[:, 0], rfm_scaled[:, 1], rfm_scaled[:, 2], c=cluster_labels, s=50, cmap='Set1', edgecolor='k')
    
    ax.grid(True)
    ax.set_xlabel('Recency')
    ax.set_ylabel('Frequency')
    ax.set_zlabel('Monetary')
    ax.set_title('3D View of Clusters')
    
    plt.show()
    return fig

def show_dashboard(data, key_suffix=''):
    st.subheader("Data Overview")
    st.write(data)

    rfm = process_rfm(data)

    scaler = StandardScaler()
    rfm_scaled = scaler.fit_transform(rfm[['Recency', 'Frequency', 'Monetary']])

    if rfm_scaled.shape[0] > 0:
        # Find optimal K
        st.subheader(f"Elbow Method Result - Optimal K{key_suffix}")
        optimal_k = find_optimal_k(rfm_scaled)
        if optimal_k is not None:
            st.write(f"Optimal number of clusters: {optimal_k}")

            cluster_labels = cluster_rfm(rfm_scaled, optimal_k)
            
            st.subheader(f"3D Clustering Visualization{key_suffix}")
            cluster_plot = plot_3d_clusters(rfm_scaled, cluster_labels)
            st.pyplot(cluster_plot)
    else:
        st.error("Tidak ada data yang valid untuk clustering.")