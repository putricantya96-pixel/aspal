import streamlit as st
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

st.set_page_config(
    page_title="GradAspal - Analisis Gradasi Agregat",
    page_icon="🏗️",
    layout="wide"
)

# Custom CSS for Modern Pavement Aesthetic
st.markdown("""
    <style>
    .main {
        background-color: #f8f9fa;
    }
    .stApp header {
        background-color: transparent;
    }
    h1, h2, h3 {
        color: #1e293b;
        font-family: 'Inter', sans-serif;
    }
    .stButton>button {
        background-color: #0f766e;
        color: white;
        border-radius: 8px;
        font-weight: bold;
    }
    .stButton>button:hover {
        background-color: #115e59;
        color: white;
    }
    </style>
""", unsafe_allow_html=True)

st.title("🏗️ GradAspal")
st.subheader("Aplikasi Analisis & Visualisasi Gradasi Agregat Campuran Aspal")
st.write("Selamat datang di GradAspal! Aplikasi ini dirancang untuk membantu evaluasi distribusi ukuran butir agregat sesuai spesifikasi standar.")

# Sidebar Controls
st.sidebar.header("⚙️ Parameter Masukan")
sieve_sizes = [37.5, 25.0, 19.0, 12.5, 9.5, 4.75, 2.36, 1.18, 0.6, 0.3, 0.15, 0.075]

st.sidebar.markdown("### Persen Lolos Ayakan (%)")
data = []
default_passing = [100.0, 95.0, 85.0, 70.0, 60.0, 45.0, 32.0, 22.0, 15.0, 10.0, 6.0, 4.0]
default_min = [100.0, 90.0, 75.0, 60.0, 50.0, 35.0, 24.0, 15.0, 9.0, 5.0, 3.0, 2.0]
default_max = [100.0, 100.0, 95.0, 80.0, 70.0, 55.0, 42.0, 30.0, 22.0, 15.0, 10.0, 7.0]

for i, size in enumerate(sieve_sizes):
    val = st.sidebar.number_input(f"Ukuran {size} mm", value=default_passing[i], min_value=0.0, max_value=100.0, step=0.5)
    data.append({
        "Ukuran Saringan (mm)": size,
        "% Lolos Sample": val,
        "Batas Min (%)": default_min[i],
        "Batas Maks (%)": default_max[i]
    })

df = pd.DataFrame(data)

col1, col2 = st.columns([1, 1])

with col1:
    st.markdown("### 📊 Tabel Analisis Gradasi")
    st.dataframe(df, use_container_width=True)

with col2:
    st.markdown("### 📈 Kurva Gradasi Agregat")
    fig, ax = plt.subplots(figsize=(8, 5))
    
    # Plot data
    ax.plot(df["Ukuran Saringan (mm)"], df["% Lolos Sample"], marker='o', color='#0f766e', linewidth=2, label='Sampel Uji')
    ax.plot(df["Ukuran Saringan (mm)"], df["Batas Min (%)"], linestyle='--', color='#ef4444', label='Batas Minimum')
    ax.plot(df["Ukuran Saringan (mm)"], df["Batas Maks (%)"], linestyle='--', color='#3b82f6', label='Batas Maksimum')
    
    ax.set_xscale('log')
    ax.set_xlabel("Ukuran Saringan (mm) - Skala Log")
    ax.set_ylabel("Persen Lolos (%)")
    ax.set_title("Kurva Analisis Ayakan Agregat")
    ax.grid(True, which="both", ls="--", alpha=0.5)
    ax.legend()
    
    st.pyplot(fig)

st.success("Analisis selesai. File siap digunakan!")
