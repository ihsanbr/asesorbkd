import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np

# Page config
st.set_page_config(
    page_title="Dashboard Monitoring Asesor BKD",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better look
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1E3A5F;
        text-align: center;
        margin-bottom: 1rem;
    }
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 1rem;
        border-radius: 10px;
        color: white;
        text-align: center;
    }
    .stMetric {
        background-color: #f0f2f6;
        padding: 15px;
        border-radius: 10px;
        border-left: 5px solid #1E88E5;
    }
    div[data-testid="stSidebar"] {
        background-color: #1E3A5F;
    }
    div[data-testid="stSidebar"] * {
        color: white !important;
    }
</style>
""", unsafe_allow_html=True)

# Sidebar
st.sidebar.title("📊 Asesor BKD")
st.sidebar.markdown("---")
menu = st.sidebar.radio(
    "Menu Navigasi",
    ["🏠 Dashboard", "📋 Daftar Asesor BKD", "⚠️ Kekurangan Asesor BKD", "✅ Eligible Asesor BKD"]
)

st.sidebar.markdown("---")
st.sidebar.subheader("📁 Upload Data Excel")

uploaded_asesor = st.sidebar.file_uploader("Data Asesor BKD", type=["xlsx", "xls"], key="asesor")
uploaded_dosen = st.sidebar.file_uploader("Data Dosen Masuk Periode BKD", type=["xlsx", "xls"], key="dosen")

# Load data
@st.cache_data
def load_data(file):
    return pd.read_excel(file)

# Default sample data if no upload (relative path)
import os
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def get_sample_asesor():
    return pd.read_excel(os.path.join(BASE_DIR, "data_asesor.xlsx"))

def get_sample_dosen():
    return pd.read_excel(os.path.join(BASE_DIR, "data_dosen.xlsx"))

if uploaded_asesor is not None:
    df_asesor = load_data(uploaded_asesor)
else:
    df_asesor = get_sample_asesor()
    st.sidebar.info("Menggunakan data sample Asesor")

if uploaded_dosen is not None:
    df_dosen = load_data(uploaded_dosen)
else:
    df_dosen = get_sample_dosen()
    st.sidebar.info("Menggunakan data sample Dosen")

# Ensure columns exist
required_asesor = ['ID Peg', 'NUPTK', 'Nama', 'Fakultas', 'Prodi']
required_dosen = ['ID Peg', 'NUPTK', 'Nama', 'Fakultas', 'Prodi', 'Jabatan Akademik', 'Status Dosen', 'Status SERDOS']

# Add Pendidikan if missing (for eligibility)
if 'Pendidikan Terakhir' not in df_dosen.columns:
    # Infer or set default for demo
    df_dosen['Pendidikan Terakhir'] = np.random.choice(['S2', 'S3'], len(df_dosen), p=[0.5, 0.5])

# ========== HELPER FUNCTIONS ==========
def calculate_kekurangan(df_asesor, df_dosen):
    asesor_count = df_asesor.groupby(['Fakultas', 'Prodi']).size().reset_index(name='Jumlah Asesor')
    dosen_count = df_dosen.groupby(['Fakultas', 'Prodi']).size().reset_index(name='Jumlah Dosen')
    
    comparison = pd.merge(dosen_count, asesor_count, on=['Fakultas', 'Prodi'], how='left')
    comparison['Jumlah Asesor'] = comparison['Jumlah Asesor'].fillna(0).astype(int)
    comparison['Ideal (30%)'] = (comparison['Jumlah Dosen'] * 0.3).round(0).astype(int)
    comparison['Kekurangan'] = comparison['Ideal (30%)'] - comparison['Jumlah Asesor']
    comparison['Kekurangan'] = comparison['Kekurangan'].apply(lambda x: max(0, x))
    comparison['Status'] = comparison.apply(
        lambda row: '✅ Cukup' if row['Jumlah Asesor'] >= row['Ideal (30%)'] else '⚠️ Kurang', axis=1
    )
    comparison['Persentase Asesor'] = (comparison['Jumlah Asesor'] / comparison['Jumlah Dosen'] * 100).round(1)
    return comparison.sort_values('Kekurangan', ascending=False)

def get_eligible(df_dosen):
    # Criteria:
    # Status Dosen: Dosen Tetap (aktif, tercatat PDDIKTI - assume Status Dosen indicates this)
    # Kualifikasi:
    # - S3 + minimal Lektor
    # - S2 + minimal Lektor Kepala
    # Serdos: Sudah
    
    def is_eligible(row):
        if row['Status Dosen'] != 'Dosen Tetap':
            return False
        if row['Status SERDOS'] != 'Sudah':
            return False
        
        jabatan = str(row['Jabatan Akademik']).lower()
        pendidikan = str(row.get('Pendidikan Terakhir', '')).upper()
        
        # Rank positions
        if 'guru besar' in jabatan or 'profesor' in jabatan:
            jab_rank = 4
        elif 'lektor kepala' in jabatan:
            jab_rank = 3
        elif 'lektor' in jabatan:
            jab_rank = 2
        elif 'asisten' in jabatan:
            jab_rank = 1
        else:
            jab_rank = 0
        
        if 'S3' in pendidikan or 'Doktor' in pendidikan:
            return jab_rank >= 2  # minimal Lektor
        elif 'S2' in pendidikan or 'Magister' in pendidikan:
            return jab_rank >= 3  # minimal Lektor Kepala
        return False
    
    eligible = df_dosen[df_dosen.apply(is_eligible, axis=1)].copy()
    return eligible

# ========== DASHBOARD ==========
if menu == "🏠 Dashboard":
    st.markdown('<p class="main-header">📊 Dashboard Monitoring Asesor BKD</p>', unsafe_allow_html=True)
    st.markdown("### Sistem Monitoring Kecukupan dan Kelayakan Asesor Beban Kerja Dosen")
    
    # Metrics
    total_asesor = len(df_asesor)
    total_dosen = len(df_dosen)
    df_kekurangan = calculate_kekurangan(df_asesor, df_dosen)
    total_kekurangan = df_kekurangan['Kekurangan'].sum()
    prodi_kurang = len(df_kekurangan[df_kekurangan['Kekurangan'] > 0])
    df_eligible = get_eligible(df_dosen)
    total_eligible = len(df_eligible)
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Total Asesor BKD", total_asesor, help="Jumlah asesor yang terdaftar")
    with col2:
        st.metric("Total Dosen Periode", total_dosen, help="Dosen yang masuk periode BKD")
    with col3:
        st.metric("Total Kekurangan", int(total_kekurangan), delta=f"{prodi_kurang} Prodi kurang", delta_color="inverse")
    with col4:
        st.metric("Eligible Asesor", total_eligible, help="Dosen yang memenuhi syarat menjadi asesor")
    
    st.markdown("---")
    
    # Graphs
    col_left, col_right = st.columns(2)
    
    with col_left:
        st.subheader("📈 Distribusi Asesor per Fakultas")
        asesor_fakultas = df_asesor.groupby('Fakultas').size().reset_index(name='Jumlah')
        fig1 = px.bar(
            asesor_fakultas, x='Fakultas', y='Jumlah',
            color='Jumlah', color_continuous_scale='Blues',
            text='Jumlah'
        )
        fig1.update_traces(textposition='outside')
        fig1.update_layout(showlegend=False, height=350, margin=dict(t=30, b=30))
        st.plotly_chart(fig1, use_container_width=True)
    
    with col_right:
        st.subheader("📊 Distribusi Asesor per Prodi")
        asesor_prodi = df_asesor.groupby('Prodi').size().reset_index(name='Jumlah').sort_values('Jumlah', ascending=True)
        fig2 = px.bar(
            asesor_prodi, x='Jumlah', y='Prodi', orientation='h',
            color='Jumlah', color_continuous_scale='Teal',
            text='Jumlah'
        )
        fig2.update_traces(textposition='outside')
        fig2.update_layout(showlegend=False, height=350, margin=dict(t=30, b=30))
        st.plotly_chart(fig2, use_container_width=True)
    
    # Kekurangan chart
    st.subheader("⚠️ Status Kecukupan Asesor per Prodi (Ideal 30%)")
    fig3 = go.Figure()
    fig3.add_trace(go.Bar(
        name='Jumlah Asesor',
        x=df_kekurangan['Prodi'],
        y=df_kekurangan['Jumlah Asesor'],
        marker_color='#2196F3'
    ))
    fig3.add_trace(go.Bar(
        name='Ideal (30%)',
        x=df_kekurangan['Prodi'],
        y=df_kekurangan['Ideal (30%)'],
        marker_color='#FF9800'
    ))
    fig3.add_trace(go.Bar(
        name='Kekurangan',
        x=df_kekurangan['Prodi'],
        y=df_kekurangan['Kekurangan'],
        marker_color='#F44336'
    ))
    fig3.update_layout(
        barmode='group',
        height=400,
        xaxis_tickangle=-45,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
    )
    st.plotly_chart(fig3, use_container_width=True)
    
    # Pie charts
    col_a, col_b = st.columns(2)
    with col_a:
        st.subheader("🥧 Status Kecukupan Prodi")
        status_count = df_kekurangan['Status'].value_counts().reset_index()
        status_count.columns = ['Status', 'Jumlah']
        fig4 = px.pie(status_count, values='Jumlah', names='Status', 
                      color='Status',
                      color_discrete_map={'✅ Cukup': '#4CAF50', '⚠️ Kurang': '#F44336'},
                      hole=0.4)
        fig4.update_layout(height=300)
        st.plotly_chart(fig4, use_container_width=True)
    
    with col_b:
        st.subheader("🥧 Eligible vs Non-Eligible")
        eligible_status = pd.DataFrame({
            'Kategori': ['Eligible', 'Belum Eligible'],
            'Jumlah': [total_eligible, total_dosen - total_eligible]
        })
        fig5 = px.pie(eligible_status, values='Jumlah', names='Kategori',
                      color='Kategori',
                      color_discrete_map={'Eligible': '#4CAF50', 'Belum Eligible': '#9E9E9E'},
                      hole=0.4)
        fig5.update_layout(height=300)
        st.plotly_chart(fig5, use_container_width=True)
    
    # Summary table
    st.subheader("📋 Ringkasan Kekurangan per Prodi")
    st.dataframe(
        df_kekurangan[['Fakultas', 'Prodi', 'Jumlah Dosen', 'Jumlah Asesor', 'Ideal (30%)', 'Kekurangan', 'Persentase Asesor', 'Status']],
        use_container_width=True,
        hide_index=True
    )

# ========== DAFTAR ASESOR ==========
elif menu == "📋 Daftar Asesor BKD":
    st.markdown('<p class="main-header">📋 Daftar Asesor BKD</p>', unsafe_allow_html=True)
    
    st.metric("Total Asesor", len(df_asesor))
    
    # Filters
    col1, col2 = st.columns(2)
    with col1:
        fakultas_filter = st.multiselect("Filter Fakultas", options=sorted(df_asesor['Fakultas'].unique()), default=None)
    with col2:
        prodi_filter = st.multiselect("Filter Prodi", options=sorted(df_asesor['Prodi'].unique()), default=None)
    
    filtered = df_asesor.copy()
    if fakultas_filter:
        filtered = filtered[filtered['Fakultas'].isin(fakultas_filter)]
    if prodi_filter:
        filtered = filtered[filtered['Prodi'].isin(prodi_filter)]
    
    st.dataframe(filtered, use_container_width=True, hide_index=True)
    
    # Download
    csv = filtered.to_csv(index=False).encode('utf-8')
    st.download_button("⬇️ Download CSV", csv, "daftar_asesor_bkd.csv", "text/csv")
    
    # Charts
    st.subheader("Visualisasi")
    c1, c2 = st.columns(2)
    with c1:
        fig = px.pie(filtered, names='Fakultas', title='Proporsi per Fakultas', hole=0.3)
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        fig = px.histogram(filtered, x='Prodi', title='Jumlah per Prodi', color='Fakultas')
        fig.update_layout(xaxis_tickangle=-45)
        st.plotly_chart(fig, use_container_width=True)

# ========== KEKURANGAN ASESOR ==========
elif menu == "⚠️ Kekurangan Asesor BKD":
    st.markdown('<p class="main-header">⚠️ Kekurangan Asesor BKD</p>', unsafe_allow_html=True)
    st.info("💡 Idealnya, setiap Prodi memiliki Asesor sebanyak **30%** dari jumlah Dosen pada Prodi tersebut.")
    
    df_kekurangan = calculate_kekurangan(df_asesor, df_dosen)
    
    # Metrics
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total Kekurangan Asesor", int(df_kekurangan['Kekurangan'].sum()))
    with col2:
        st.metric("Prodi yang Kurang", len(df_kekurangan[df_kekurangan['Kekurangan'] > 0]))
    with col3:
        st.metric("Prodi yang Cukup", len(df_kekurangan[df_kekurangan['Kekurangan'] == 0]))
    
    # Filter
    status_filter = st.selectbox("Filter Status", ["Semua", "✅ Cukup", "⚠️ Kurang"])
    if status_filter != "Semua":
        df_show = df_kekurangan[df_kekurangan['Status'] == status_filter]
    else:
        df_show = df_kekurangan
    
    st.dataframe(
        df_show.style.apply(
            lambda row: ['background-color: #FFCDD2' if row['Status'] == '⚠️ Kurang' else 'background-color: #C8E6C9'] * len(row),
            axis=1
        ),
        use_container_width=True,
        hide_index=True
    )
    
    # Chart
    st.subheader("Perbandingan Jumlah Asesor vs Ideal")
    fig = go.Figure()
    fig.add_trace(go.Bar(name='Jumlah Asesor', x=df_kekurangan['Prodi'], y=df_kekurangan['Jumlah Asesor'], marker_color='#2196F3'))
    fig.add_trace(go.Scatter(name='Ideal 30%', x=df_kekurangan['Prodi'], y=df_kekurangan['Ideal (30%)'], mode='lines+markers', line=dict(color='#FF9800', width=3)))
    fig.update_layout(height=400, xaxis_tickangle=-45, barmode='group')
    st.plotly_chart(fig, use_container_width=True)
    
    # Detail per prodi yang kurang
    kurang = df_kekurangan[df_kekurangan['Kekurangan'] > 0]
    if not kurang.empty:
        st.subheader("Detail Prodi yang Mengalami Kekurangan")
        for _, row in kurang.iterrows():
            with st.expander(f"{row['Prodi']} ({row['Fakultas']}) - Kurang {int(row['Kekurangan'])} Asesor"):
                st.write(f"- Jumlah Dosen: **{row['Jumlah Dosen']}**")
                st.write(f"- Jumlah Asesor saat ini: **{row['Jumlah Asesor']}**")
                st.write(f"- Ideal (30%): **{row['Ideal (30%)']}**")
                st.write(f"- Persentase Asesor: **{row['Persentase Asesor']}%**")
                st.progress(min(row['Persentase Asesor'] / 30, 1.0))

# ========== ELIGIBLE ASESOR ==========
elif menu == "✅ Eligible Asesor BKD":
    st.markdown('<p class="main-header">✅ Eligible Asesor BKD</p>', unsafe_allow_html=True)
    
    st.markdown("""
    ### Kriteria Eligible Asesor BKD:
    1. **Status Dosen**: Dosen Tetap yang masih aktif dan tercatat di PDDIKTI.
    2. **Kualifikasi Akademik dan Jabatan Fungsional**:
       - Minimal berijazah **Doktor (S3)** dengan jabatan akademik minimal **Lektor**.
       - Atau minimal berijazah **Magister (S2)** dengan jabatan akademik minimal **Lektor Kepala**.
    3. **Sertifikasi Dosen**: Memiliki **Sertifikat Pendidik untuk Dosen (Serdos)**.
    """)
    
    df_eligible = get_eligible(df_dosen)
    
    col1, col2 = st.columns(2)
    with col1:
        st.metric("Total Eligible", len(df_eligible))
    with col2:
        st.metric("% dari Total Dosen", f"{len(df_eligible)/len(df_dosen)*100:.1f}%")
    
    # Filters
    c1, c2, c3 = st.columns(3)
    with c1:
        fak_filter = st.multiselect("Fakultas", sorted(df_eligible['Fakultas'].unique()) if not df_eligible.empty else [])
    with c2:
        prodi_f = st.multiselect("Prodi", sorted(df_eligible['Prodi'].unique()) if not df_eligible.empty else [])
    with c3:
        jab_filter = st.multiselect("Jabatan", sorted(df_eligible['Jabatan Akademik'].unique()) if not df_eligible.empty else [])
    
    filtered_el = df_eligible.copy()
    if fak_filter:
        filtered_el = filtered_el[filtered_el['Fakultas'].isin(fak_filter)]
    if prodi_f:
        filtered_el = filtered_el[filtered_el['Prodi'].isin(prodi_f)]
    if jab_filter:
        filtered_el = filtered_el[filtered_el['Jabatan Akademik'].isin(jab_filter)]
    
    display_cols = ['ID Peg', 'NUPTK', 'Nama', 'Fakultas', 'Prodi', 'Jabatan Akademik', 'Status Dosen', 'Status SERDOS', 'Pendidikan Terakhir']
    display_cols = [c for c in display_cols if c in filtered_el.columns]
    st.dataframe(filtered_el[display_cols], use_container_width=True, hide_index=True)
    
    csv = filtered_el.to_csv(index=False).encode('utf-8')
    st.download_button("⬇️ Download Eligible CSV", csv, "eligible_asesor_bkd.csv", "text/csv")
    
    # Charts
    if not df_eligible.empty:
        st.subheader("Visualisasi Eligible")
        c1, c2 = st.columns(2)
        with c1:
            fig = px.bar(df_eligible.groupby('Fakultas').size().reset_index(name='Jumlah'), 
                         x='Fakultas', y='Jumlah', color='Jumlah', title='Eligible per Fakultas',
                         color_continuous_scale='Greens')
            st.plotly_chart(fig, use_container_width=True)
        with c2:
            fig = px.pie(df_eligible, names='Jabatan Akademik', title='Distribusi Jabatan Eligible', hole=0.3)
            st.plotly_chart(fig, use_container_width=True)
        
        # Comparison: already assessor vs eligible
        st.subheader("Potensi Rekrutmen")
        # Match by ID Peg or Nama if possible
        already = set(df_asesor['ID Peg'].astype(str))
        df_eligible['Sudah Asesor'] = df_eligible['ID Peg'].astype(str).isin(already)
        potensi = df_eligible[~df_eligible['Sudah Asesor']]
        st.write(f"Dari {len(df_eligible)} eligible, **{len(potensi)}** belum menjadi Asesor (potensi rekrutmen).")
        
        if not potensi.empty:
            st.dataframe(potensi[display_cols], use_container_width=True, hide_index=True)

st.sidebar.markdown("---")
st.sidebar.caption("Dashboard Monitoring Asesor BKD v1.0")
st.sidebar.caption("Data diimport dari Excel")
