import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
import os
from io import BytesIO
from datetime import datetime

# ==================== PAGE CONFIG ====================
st.set_page_config(
    page_title="Dashboard Monitoring Asesor BKD | UNS",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==================== THEME - PEAR UNS (#0e7da7) ====================
PRIMARY = "#0e7da7"
PRIMARY_DARK = "#0a5f80"
PRIMARY_LIGHT = "#e6f4f9"
ACCENT = "#1a9bc7"
SIDEBAR_BG = "#0a4d6b"

st.markdown(f"""
<style>
    .main-header {{
        font-size: 1.9rem;
        font-weight: 700;
        color: {PRIMARY_DARK};
        margin: 0;
        line-height: 1.2;
    }}
    .sub-header {{
        font-size: 0.95rem;
        color: #555;
        margin-bottom: 1rem;
    }}
    div[data-testid="stMetric"] {{
        background: linear-gradient(135deg, {PRIMARY} 0%, {PRIMARY_DARK} 100%);
        padding: 16px 20px;
        border-radius: 12px;
        color: white !important;
        box-shadow: 0 4px 12px rgba(14,125,167,0.25);
    }}
    div[data-testid="stMetric"] label {{
        color: rgba(255,255,255,0.9) !important;
    }}
    div[data-testid="stMetric"] [data-testid="stMetricValue"] {{
        color: white !important;
        font-size: 1.8rem !important;
    }}
    div[data-testid="stMetric"] [data-testid="stMetricDelta"] {{
        color: #c8f0ff !important;
    }}

    /* ===== SIDEBAR ===== */
    section[data-testid="stSidebar"] {{
        background: linear-gradient(180deg, {SIDEBAR_BG} 0%, {PRIMARY_DARK} 100%) !important;
    }}
    /* Teks navigasi & judul sidebar: putih */
    section[data-testid="stSidebar"] h1,
    section[data-testid="stSidebar"] h2,
    section[data-testid="stSidebar"] h3,
    section[data-testid="stSidebar"] .stMarkdown,
    section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"],
    section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p,
    section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] span,
    section[data-testid="stSidebar"] label,
    section[data-testid="stSidebar"] .stRadio label,
    section[data-testid="stSidebar"] .stCaption,
    section[data-testid="stSidebar"] small {{
        color: #ffffff !important;
    }}
    section[data-testid="stSidebar"] hr {{
        border-color: rgba(255,255,255,0.25) !important;
    }}

    /* Label file uploader: putih tebal agar terbaca di sidebar gelap */
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] label,
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] label span,
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] p {{
        color: #ffffff !important;
        font-weight: 600 !important;
        font-size: 0.95rem !important;
    }}

    /* Kotak dropzone upload: putih + teks gelap + border biru */
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] section,
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] section > div,
    section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] {{
        background-color: #ffffff !important;
        border: 2px dashed {PRIMARY} !important;
        border-radius: 10px !important;
        color: {PRIMARY_DARK} !important;
    }}
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] section *,
    section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] *,
    section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] span,
    section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] small,
    section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] svg {{
        color: {PRIMARY_DARK} !important;
        fill: {PRIMARY_DARK} !important;
    }}
    /* Tombol Browse di dalam dropzone */
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] button {{
        background-color: {PRIMARY} !important;
        color: #ffffff !important;
        border: none !important;
        border-radius: 8px !important;
        font-weight: 600 !important;
    }}
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] button:hover {{
        background-color: {PRIMARY_DARK} !important;
        color: #ffffff !important;
    }}
    /* Nama file yang sudah di-upload */
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] [data-testid="stFileUploaderFileName"],
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] [data-testid="stFileUploaderFileData"] {{
        color: #ffffff !important;
    }}

    /* Info / caption di sidebar */
    section[data-testid="stSidebar"] [data-testid="stAlert"] {{
        background-color: rgba(255,255,255,0.12) !important;
        color: #ffffff !important;
    }}
    section[data-testid="stSidebar"] [data-testid="stAlert"] * {{
        color: #ffffff !important;
    }}

    /* Tombol utama di area konten */
    .stDownloadButton > button, .stButton > button {{
        background-color: {PRIMARY} !important;
        color: white !important;
        border: none !important;
        border-radius: 8px !important;
        font-weight: 600 !important;
    }}
    .stDownloadButton > button:hover, .stButton > button:hover {{
        background-color: {PRIMARY_DARK} !important;
    }}
    .stExpander {{
        border-left: 4px solid {PRIMARY};
    }}
    .stAlert {{
        border-left-color: {PRIMARY} !important;
    }}
</style>
""", unsafe_allow_html=True)

# ==================== PATHS ====================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOGO_PATH = os.path.join(BASE_DIR, "assets", "logo_uns.png")

# ==================== HELPERS ====================
def to_excel_bytes(df: pd.DataFrame, sheet_name: str = "Sheet1") -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return buffer.getvalue()

def load_excel(file):
    return pd.read_excel(file)

def normalize_asesor(df):
    """Pastikan kolom asesor konsisten & aman untuk data_editor."""
    expected = ["ID Peg", "NUPTK", "Nama", "Fakultas", "Prodi"]
    df = df.copy()
    for c in expected:
        if c not in df.columns:
            df[c] = ""
    df = df[expected].fillna("").astype(str)
    return df.reset_index(drop=True)

def get_sample_asesor():
    return normalize_asesor(pd.read_excel(os.path.join(BASE_DIR, "data_asesor.xlsx")))

def get_sample_dosen():
    return pd.read_excel(os.path.join(BASE_DIR, "data_dosen.xlsx"))

def normalize_text(val):
    return str(val if val is not None else "").strip().lower()

def calculate_kekurangan(df_asesor, df_dosen):
    """Perbandingan per Fakultas+Prodi: Ideal = 30% jumlah dosen di prodi yang sama."""
    a = df_asesor.copy()
    d = df_dosen.copy()
    for col in ["Fakultas", "Prodi"]:
        if col in a.columns:
            a[col] = a[col].astype(str).str.strip()
        if col in d.columns:
            d[col] = d[col].astype(str).str.strip()

    asesor_count = a.groupby(["Fakultas", "Prodi"], dropna=False).size().reset_index(name="Jumlah Asesor")
    dosen_count = d.groupby(["Fakultas", "Prodi"], dropna=False).size().reset_index(name="Jumlah Dosen")

    # Outer agar prodi yang hanya punya asesor atau hanya dosen tetap tampil
    comparison = pd.merge(dosen_count, asesor_count, on=["Fakultas", "Prodi"], how="outer")
    comparison["Jumlah Dosen"] = comparison["Jumlah Dosen"].fillna(0).astype(int)
    comparison["Jumlah Asesor"] = comparison["Jumlah Asesor"].fillna(0).astype(int)
    comparison["Ideal (30%)"] = (comparison["Jumlah Dosen"] * 0.3).round(0).astype(int)
    comparison["Kekurangan"] = (comparison["Ideal (30%)"] - comparison["Jumlah Asesor"]).clip(lower=0)
    comparison["Kelebihan"] = (comparison["Jumlah Asesor"] - comparison["Ideal (30%)"]).clip(lower=0)
    comparison["Status"] = comparison.apply(
        lambda r: (
            "➖ Tidak ada Dosen" if r["Jumlah Dosen"] == 0
            else ("✅ Cukup" if r["Jumlah Asesor"] >= r["Ideal (30%)"] else "⚠️ Kurang")
        ),
        axis=1,
    )
    comparison["Persentase Asesor"] = comparison.apply(
        lambda r: round(r["Jumlah Asesor"] / r["Jumlah Dosen"] * 100, 1) if r["Jumlah Dosen"] > 0 else 0.0,
        axis=1,
    )
    return comparison.sort_values(["Kekurangan", "Fakultas", "Prodi"], ascending=[False, True, True])

def jabatan_rank(jabatan):
    j = normalize_text(jabatan)
    if "guru besar" in j or "profesor" in j:
        return 4
    if "lektor kepala" in j:
        return 3
    if "lektor" in j:  # Lektor (bukan Lektor Kepala)
        return 2
    if "asisten" in j:
        return 1
    return 0

def is_dosen_tetap(status):
    s = normalize_text(status)
    # Terima variasi: Dosen Tetap, Tetap, PNS, DTY, dll.
    if any(x in s for x in ["tidak tetap", "non tetap", "honorer", "luar biasa", "lb"]):
        return False
    return any(x in s for x in ["tetap", "pns", "dty", "permanen", "aktif"])

def is_serdos_sudah(status):
    s = normalize_text(status)
    if any(x in s for x in ["belum", "tidak", "no", "0", "false"]):
        return False
    return any(x in s for x in ["sudah", "yes", "1", "true", "sertifikat", "punya", "ada"])

def is_s3(pendidikan):
    p = normalize_text(pendidikan)
    return any(x in p for x in ["s3", "doktor", "doctor", "phd", "dr."])

def is_s2(pendidikan):
    p = normalize_text(pendidikan)
    return any(x in p for x in ["s2", "magister", "master", "m."])

def ensure_pendidikan(df_dosen):
    """Jika kolom Pendidikan Terakhir tidak ada, isi default wajar dari jabatan."""
    df = df_dosen.copy()
    if "Pendidikan Terakhir" not in df.columns:
        def guess(row):
            r = jabatan_rank(row.get("Jabatan Akademik", ""))
            if r >= 3:
                return "S3"
            if r == 2:
                return "S3"  # Lektor sering S3
            return "S2"
        df["Pendidikan Terakhir"] = df.apply(guess, axis=1)
    else:
        df["Pendidikan Terakhir"] = df["Pendidikan Terakhir"].fillna("").astype(str)
    return df

def get_eligible(df_dosen):
    """
    Eligible jika:
    1. Status Dosen tetap/aktif
    2. S3 + minimal Lektor  ATAU  S2 + minimal Lektor Kepala
    3. Serdos sudah
    """
    df = ensure_pendidikan(df_dosen)

    def is_eligible(row):
        if not is_dosen_tetap(row.get("Status Dosen", "")):
            return False
        if not is_serdos_sudah(row.get("Status SERDOS", "")):
            return False
        rank = jabatan_rank(row.get("Jabatan Akademik", ""))
        pend = row.get("Pendidikan Terakhir", "")
        if is_s3(pend):
            return rank >= 2  # minimal Lektor
        if is_s2(pend):
            return rank >= 3  # minimal Lektor Kepala
        # Jika pendidikan kosong tapi Guru Besar / Lektor Kepala → anggap eligible (asumsi S3)
        if rank >= 3:
            return True
        return False

    result = df[df.apply(is_eligible, axis=1)].copy()
    return result.reset_index(drop=True)

def render_header(title: str, subtitle: str = ""):
    col_logo, col_text = st.columns([1, 6])
    with col_logo:
        if os.path.exists(LOGO_PATH):
            st.image(LOGO_PATH, width=72)
        else:
            st.markdown("🎓")
    with col_text:
        st.markdown(f'<p class="main-header">{title}</p>', unsafe_allow_html=True)
        if subtitle:
            st.markdown(f'<p class="sub-header">{subtitle}</p>', unsafe_allow_html=True)

# ==================== SIDEBAR ====================
st.sidebar.markdown("### Asesor BKD UNS")
if os.path.exists(LOGO_PATH):
    st.sidebar.image(LOGO_PATH, width=120)
st.sidebar.markdown("---")

menu = st.sidebar.radio(
    "Menu Navigasi",
    [
        "Dashboard",
        "Daftar Asesor BKD",
        "Data Dosen",
        "Kekurangan Asesor BKD",
        "Eligible Asesor BKD",
    ],
)

st.sidebar.markdown("---")
st.sidebar.subheader("📁 Upload Data Excel")
uploaded_asesor = st.sidebar.file_uploader("Data Asesor BKD", type=["xlsx", "xls"], key="up_asesor")
uploaded_dosen = st.sidebar.file_uploader("Data Dosen Masuk Periode BKD", type=["xlsx", "xls"], key="up_dosen")

# ==================== SESSION STATE ====================
if "df_asesor" not in st.session_state:
    if uploaded_asesor is not None:
        st.session_state.df_asesor = normalize_asesor(load_excel(uploaded_asesor))
    else:
        st.session_state.df_asesor = get_sample_asesor()
        st.sidebar.info("📦 Menggunakan data sample Asesor")

if uploaded_asesor is not None:
    if st.session_state.get("_last_asesor_name") != uploaded_asesor.name:
        st.session_state.df_asesor = normalize_asesor(load_excel(uploaded_asesor))
        st.session_state._last_asesor_name = uploaded_asesor.name


if "df_dosen" not in st.session_state:
    if uploaded_dosen is not None:
        st.session_state.df_dosen = load_excel(uploaded_dosen)
    else:
        st.session_state.df_dosen = get_sample_dosen()
        st.sidebar.info("📦 Menggunakan data sample Dosen")

if uploaded_dosen is not None:
    if st.session_state.get("_last_dosen_name") != uploaded_dosen.name:
        st.session_state.df_dosen = load_excel(uploaded_dosen)
        st.session_state._last_dosen_name = uploaded_dosen.name

df_asesor = st.session_state.df_asesor.copy()
df_dosen = ensure_pendidikan(st.session_state.df_dosen.copy())
st.session_state.df_dosen = df_dosen

# ==================== DASHBOARD ====================
if menu == "Dashboard":
    render_header(
        "Dashboard Monitoring Asesor BKD",
        "Universitas Sebelas Maret · Sistem Monitoring Kecukupan & Kelayakan Asesor"
    )

    total_asesor = len(df_asesor)
    total_dosen = len(df_dosen)
    df_kekurangan = calculate_kekurangan(df_asesor, df_dosen)
    total_kekurangan = int(df_kekurangan["Kekurangan"].sum())
    prodi_kurang = len(df_kekurangan[df_kekurangan["Kekurangan"] > 0])
    df_eligible = get_eligible(df_dosen)
    total_eligible = len(df_eligible)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Asesor BKD", total_asesor)
    c2.metric("Total Dosen Periode", total_dosen)
    c3.metric("Total Kekurangan", total_kekurangan, delta=f"{prodi_kurang} Prodi kurang", delta_color="inverse")
    c4.metric("Eligible Asesor", total_eligible)

    st.markdown("---")

    col_l, col_r = st.columns(2)
    with col_l:
        st.subheader("📈 Distribusi Asesor per Fakultas")
        af = df_asesor.groupby("Fakultas").size().reset_index(name="Jumlah")
        fig = px.bar(af, x="Fakultas", y="Jumlah", color="Jumlah",
                     color_continuous_scale=["#a8d5e5", PRIMARY], text="Jumlah")
        fig.update_traces(textposition="outside")
        fig.update_layout(showlegend=False, height=340, margin=dict(t=20, b=20),
                          plot_bgcolor="white", paper_bgcolor="white")
        st.plotly_chart(fig, use_container_width=True)

    with col_r:
        st.subheader("📊 Distribusi Asesor per Prodi")
        ap = df_asesor.groupby("Prodi").size().reset_index(name="Jumlah").sort_values("Jumlah")
        fig = px.bar(ap, x="Jumlah", y="Prodi", orientation="h", color="Jumlah",
                     color_continuous_scale=["#a8d5e5", PRIMARY], text="Jumlah")
        fig.update_traces(textposition="outside")
        fig.update_layout(showlegend=False, height=340, margin=dict(t=20, b=20),
                          plot_bgcolor="white", paper_bgcolor="white")
        st.plotly_chart(fig, use_container_width=True)

    st.subheader("⚠️ Status Kecukupan Asesor per Prodi (Ideal 30%)")
    fig = go.Figure()
    fig.add_trace(go.Bar(name="Jumlah Asesor", x=df_kekurangan["Prodi"],
                         y=df_kekurangan["Jumlah Asesor"], marker_color=PRIMARY))
    fig.add_trace(go.Bar(name="Ideal (30%)", x=df_kekurangan["Prodi"],
                         y=df_kekurangan["Ideal (30%)"], marker_color="#f5a623"))
    fig.add_trace(go.Bar(name="Kekurangan", x=df_kekurangan["Prodi"],
                         y=df_kekurangan["Kekurangan"], marker_color="#e74c3c"))
    fig.update_layout(barmode="group", height=400, xaxis_tickangle=-40,
                      legend=dict(orientation="h", y=1.08),
                      plot_bgcolor="white", paper_bgcolor="white")
    st.plotly_chart(fig, use_container_width=True)

    ca, cb = st.columns(2)
    with ca:
        st.subheader("🥧 Status Kecukupan Prodi")
        sc = df_kekurangan["Status"].value_counts().reset_index()
        sc.columns = ["Status", "Jumlah"]
        fig = px.pie(sc, values="Jumlah", names="Status", hole=0.45,
                     color="Status",
                     color_discrete_map={"✅ Cukup": "#27ae60", "⚠️ Kurang": "#e74c3c"})
        fig.update_layout(height=300, margin=dict(t=10, b=10))
        st.plotly_chart(fig, use_container_width=True)
    with cb:
        st.subheader("🥧 Eligible vs Non-Eligible")
        es = pd.DataFrame({
            "Kategori": ["Eligible", "Belum Eligible"],
            "Jumlah": [total_eligible, total_dosen - total_eligible]
        })
        fig = px.pie(es, values="Jumlah", names="Kategori", hole=0.45,
                     color="Kategori",
                     color_discrete_map={"Eligible": PRIMARY, "Belum Eligible": "#b0bec5"})
        fig.update_layout(height=300, margin=dict(t=10, b=10))
        st.plotly_chart(fig, use_container_width=True)

    st.subheader("📋 Ringkasan Kekurangan per Prodi")
    st.dataframe(
        df_kekurangan[["Fakultas", "Prodi", "Jumlah Dosen", "Jumlah Asesor",
                       "Ideal (30%)", "Kekurangan", "Persentase Asesor", "Status"]],
        use_container_width=True, hide_index=True
    )

# ==================== DAFTAR ASESOR ====================
elif menu == "Daftar Asesor BKD":
    render_header("Daftar Asesor BKD", "Kelola data asesor: edit, pilih, hapus, dan export Excel")

    st.metric("Total Asesor", len(st.session_state.df_asesor))

    fc1, fc2, fc3 = st.columns([2, 2, 1])
    with fc1:
        fakultas_opts = sorted(st.session_state.df_asesor["Fakultas"].dropna().unique())
        f_filter = st.multiselect("Filter Fakultas", fakultas_opts, key="as_fak")
    with fc2:
        prodi_opts = sorted(st.session_state.df_asesor["Prodi"].dropna().unique())
        p_filter = st.multiselect("Filter Prodi", prodi_opts, key="as_pro")
    with fc3:
        page_size = st.selectbox("Baris/halaman", [10, 20, 50, 100], index=1, key="as_ps")

    df_full = normalize_asesor(st.session_state.df_asesor)
    if f_filter:
        df_full = df_full[df_full["Fakultas"].isin(f_filter)]
    if p_filter:
        df_full = df_full[df_full["Prodi"].isin(p_filter)]
    df_full = df_full.reset_index(drop=True)

    total_rows = len(df_full)
    total_pages = max(1, (total_rows + page_size - 1) // page_size)

    pc1, pc2, pc3 = st.columns([1, 2, 1])
    with pc1:
        if st.button("← Prev", key="as_prev") and st.session_state.get("as_page", 1) > 1:
            st.session_state.as_page = st.session_state.get("as_page", 1) - 1
            st.rerun()
    with pc2:
        page = st.number_input(
            f"Halaman (1–{total_pages})",
            min_value=1,
            max_value=total_pages,
            value=min(st.session_state.get("as_page", 1), total_pages),
            key="as_page_input",
        )
        st.session_state.as_page = int(page)
    with pc3:
        if st.button("Next →", key="as_next") and st.session_state.get("as_page", 1) < total_pages:
            st.session_state.as_page = st.session_state.get("as_page", 1) + 1
            st.rerun()

    page = int(st.session_state.get("as_page", 1))
    start_i = (page - 1) * page_size
    end_i = min(start_i + page_size, total_rows)
    st.caption(f"Menampilkan baris {start_i + 1 if total_rows else 0}–{end_i} dari {total_rows} data")

    # Tambah kolom Pilih untuk delete
    df_page = df_full.iloc[start_i:end_i].copy().reset_index(drop=True)
    df_page.insert(0, "Pilih", False)

    st.markdown("### Tabel Data Asesor")
    st.caption("Centang kolom **Pilih** lalu Hapus Terpilih · Edit sel lalu Simpan Perubahan · Pagination di atas.")

    edited = st.data_editor(
        df_page,
        num_rows="fixed",
        use_container_width=True,
        hide_index=True,
        key=f"editor_asesor_p{page}",
        column_config={
            "Pilih": st.column_config.CheckboxColumn("Pilih", default=False, width="small"),
            "ID Peg": st.column_config.TextColumn("ID Peg", width="medium"),
            "NUPTK": st.column_config.TextColumn("NUPTK", width="medium"),
            "Nama": st.column_config.TextColumn("Nama", width="large"),
            "Fakultas": st.column_config.TextColumn("Fakultas", width="medium"),
            "Prodi": st.column_config.TextColumn("Prodi", width="medium"),
        },
        disabled=[],
    )

    b1, b2, b3, b4, b5 = st.columns(5)
    with b1:
        if st.button("Simpan Perubahan", use_container_width=True, key="as_save"):
            try:
                ed = edited.drop(columns=["Pilih"], errors="ignore")
                ed = normalize_asesor(ed)
                # Update baris halaman ini ke df_full lalu ke session
                full = normalize_asesor(st.session_state.df_asesor)
                # Ganti berdasarkan posisi filter+page: rebuild filtered slice
                mask_filter = pd.Series([True] * len(full))
                if f_filter:
                    mask_filter &= full["Fakultas"].isin(f_filter)
                if p_filter:
                    mask_filter &= full["Prodi"].isin(p_filter)
                idx_filtered = full.index[mask_filter].tolist()
                page_idx = idx_filtered[start_i:end_i]
                for i, src_idx in enumerate(page_idx):
                    if i < len(ed):
                        for col in ["ID Peg", "NUPTK", "Nama", "Fakultas", "Prodi"]:
                            full.at[src_idx, col] = ed.iloc[i][col]
                st.session_state.df_asesor = normalize_asesor(full)
                st.success("Perubahan halaman ini disimpan.")
                st.rerun()
            except Exception as e:
                st.error(f"Gagal menyimpan: {e}")

    with b2:
        if st.button("Hapus Terpilih", use_container_width=True, key="as_del_sel"):
            selected_ids = edited.loc[edited["Pilih"] == True, "ID Peg"].astype(str).tolist()
            if not selected_ids:
                st.warning("Tidak ada baris yang dipilih.")
            else:
                st.session_state.df_asesor = normalize_asesor(
                    st.session_state.df_asesor[
                        ~st.session_state.df_asesor["ID Peg"].astype(str).isin(selected_ids)
                    ]
                )
                st.success(f"{len(selected_ids)} data dihapus.")
                st.rerun()

    with b3:
        if st.button("Delete All", use_container_width=True, key="as_del_all", type="primary"):
            st.session_state["_confirm_del_all_asesor"] = True

    with b4:
        del_id = st.text_input("Hapus by ID Peg", placeholder="ID Peg", label_visibility="collapsed", key="as_del_id")
        if st.button("Hapus by ID", use_container_width=True, key="as_del_id_btn") and del_id:
            before = len(st.session_state.df_asesor)
            st.session_state.df_asesor = normalize_asesor(
                st.session_state.df_asesor[
                    st.session_state.df_asesor["ID Peg"].astype(str) != str(del_id).strip()
                ]
            )
            if len(st.session_state.df_asesor) < before:
                st.success(f"ID Peg {del_id} dihapus.")
                st.rerun()
            else:
                st.error("ID Peg tidak ditemukan.")

    with b5:
        ts = datetime.now().strftime("%Y%m%d_%H%M")
        excel_bytes = to_excel_bytes(normalize_asesor(st.session_state.df_asesor), "Daftar Asesor")
        st.download_button(
            "Export Excel",
            data=excel_bytes,
            file_name=f"Daftar_Asesor_BKD_{ts}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    if st.session_state.get("_confirm_del_all_asesor"):
        st.warning("Yakin hapus **SEMUA** data Asesor?")
        c_yes, c_no = st.columns(2)
        with c_yes:
            if st.button("Ya, hapus semua Asesor", key="as_del_all_yes"):
                st.session_state.df_asesor = normalize_asesor(
                    pd.DataFrame(columns=["ID Peg", "NUPTK", "Nama", "Fakultas", "Prodi"])
                )
                st.session_state._confirm_del_all_asesor = False
                st.success("Semua data Asesor dihapus.")
                st.rerun()
        with c_no:
            if st.button("Batal", key="as_del_all_no"):
                st.session_state._confirm_del_all_asesor = False
                st.rerun()

    st.markdown("---")
    st.subheader("Tambah Asesor Baru")
    with st.form("form_tambah_asesor", clear_on_submit=True):
        tc1, tc2, tc3 = st.columns(3)
        with tc1:
            new_id = st.text_input("ID Peg *")
            new_nuptk = st.text_input("NUPTK")
        with tc2:
            new_nama = st.text_input("Nama *")
            new_fak = st.text_input("Fakultas")
        with tc3:
            new_prodi = st.text_input("Prodi")
        submitted = st.form_submit_button("Tambah Asesor")
        if submitted:
            if not new_id or not new_nama:
                st.error("ID Peg dan Nama wajib diisi.")
            elif str(new_id) in st.session_state.df_asesor["ID Peg"].astype(str).values:
                st.error("ID Peg sudah ada.")
            else:
                new_row = pd.DataFrame([{
                    "ID Peg": str(new_id).strip(),
                    "NUPTK": str(new_nuptk).strip(),
                    "Nama": str(new_nama).strip(),
                    "Fakultas": str(new_fak).strip(),
                    "Prodi": str(new_prodi).strip(),
                }])
                st.session_state.df_asesor = normalize_asesor(
                    pd.concat([st.session_state.df_asesor, new_row], ignore_index=True)
                )
                st.success(f"Asesor {new_nama} ditambahkan.")
                st.rerun()

# ==================== DATA DOSEN ====================
elif menu == "Data Dosen":
    render_header("Data Dosen Masuk Periode BKD", "Daftar dosen yang masuk periode penilaian BKD — edit, pilih, hapus, pagination")

    # Pastikan session dosen editable
    if "df_dosen" not in st.session_state:
        st.session_state.df_dosen = df_dosen.copy()

    st.metric("Total Dosen Periode", len(st.session_state.df_dosen))

    d1, d2, d3, d4 = st.columns([2, 2, 2, 1])
    with d1:
        fak_opts = sorted(st.session_state.df_dosen["Fakultas"].dropna().astype(str).unique()) if "Fakultas" in st.session_state.df_dosen.columns else []
        f_f = st.multiselect("Filter Fakultas", fak_opts, key="dosen_fak")
    with d2:
        pro_opts = sorted(st.session_state.df_dosen["Prodi"].dropna().astype(str).unique()) if "Prodi" in st.session_state.df_dosen.columns else []
        p_f = st.multiselect("Filter Prodi", pro_opts, key="dosen_pro")
    with d3:
        if "Jabatan Akademik" in st.session_state.df_dosen.columns:
            jab_opts = sorted(st.session_state.df_dosen["Jabatan Akademik"].dropna().astype(str).unique())
            j_f = st.multiselect("Filter Jabatan", jab_opts, key="dosen_jab")
        else:
            j_f = []
    with d4:
        page_size = st.selectbox("Baris/halaman", [10, 20, 50, 100], index=1, key="dosen_ps")

    df_full = ensure_pendidikan(st.session_state.df_dosen.copy())
    # normalisasi string kolom kunci
    for c in df_full.columns:
        if df_full[c].dtype == object:
            df_full[c] = df_full[c].fillna("").astype(str)

    if f_f:
        df_full = df_full[df_full["Fakultas"].isin(f_f)]
    if p_f:
        df_full = df_full[df_full["Prodi"].isin(p_f)]
    if j_f and "Jabatan Akademik" in df_full.columns:
        df_full = df_full[df_full["Jabatan Akademik"].isin(j_f)]
    df_full = df_full.reset_index(drop=True)

    display_cols = [c for c in [
        "ID Peg", "NUPTK", "Nama", "Fakultas", "Prodi",
        "Jabatan Akademik", "Status Dosen", "Status SERDOS", "Pendidikan Terakhir"
    ] if c in df_full.columns]

    total_rows = len(df_full)
    total_pages = max(1, (total_rows + page_size - 1) // page_size)

    pc1, pc2, pc3 = st.columns([1, 2, 1])
    with pc1:
        if st.button("← Prev", key="dosen_prev") and st.session_state.get("dosen_page", 1) > 1:
            st.session_state.dosen_page = st.session_state.get("dosen_page", 1) - 1
            st.rerun()
    with pc2:
        page = st.number_input(
            f"Halaman (1–{total_pages})",
            min_value=1,
            max_value=total_pages,
            value=min(st.session_state.get("dosen_page", 1), total_pages),
            key="dosen_page_input",
        )
        st.session_state.dosen_page = int(page)
    with pc3:
        if st.button("Next →", key="dosen_next") and st.session_state.get("dosen_page", 1) < total_pages:
            st.session_state.dosen_page = st.session_state.get("dosen_page", 1) + 1
            st.rerun()

    page = int(st.session_state.get("dosen_page", 1))
    start_i = (page - 1) * page_size
    end_i = min(start_i + page_size, total_rows)
    st.caption(f"Menampilkan baris {start_i + 1 if total_rows else 0}–{end_i} dari {total_rows} data")

    df_page = df_full[display_cols].iloc[start_i:end_i].copy().reset_index(drop=True)
    df_page.insert(0, "Pilih", False)

    st.markdown("### Tabel Data Dosen")
    st.caption("Centang **Pilih** lalu Hapus Terpilih · Edit sel lalu Simpan · Pagination di atas.")

    col_cfg = {
        "Pilih": st.column_config.CheckboxColumn("Pilih", default=False, width="small"),
    }
    for c in display_cols:
        col_cfg[c] = st.column_config.TextColumn(c)

    edited = st.data_editor(
        df_page,
        num_rows="fixed",
        use_container_width=True,
        hide_index=True,
        key=f"editor_dosen_p{page}",
        column_config=col_cfg,
    )

    b1, b2, b3, b4 = st.columns(4)
    with b1:
        if st.button("Simpan Perubahan", use_container_width=True, key="dosen_save"):
            try:
                ed = edited.drop(columns=["Pilih"], errors="ignore")
                full = ensure_pendidikan(st.session_state.df_dosen.copy())
                for c in full.columns:
                    if full[c].dtype == object:
                        full[c] = full[c].fillna("").astype(str)
                mask = pd.Series([True] * len(full))
                if f_f:
                    mask &= full["Fakultas"].isin(f_f)
                if p_f:
                    mask &= full["Prodi"].isin(p_f)
                if j_f and "Jabatan Akademik" in full.columns:
                    mask &= full["Jabatan Akademik"].isin(j_f)
                idx_filtered = full.index[mask].tolist()
                page_idx = idx_filtered[start_i:end_i]
                for i, src_idx in enumerate(page_idx):
                    if i < len(ed):
                        for col in display_cols:
                            if col in full.columns and col in ed.columns:
                                full.at[src_idx, col] = ed.iloc[i][col]
                st.session_state.df_dosen = full.reset_index(drop=True)
                st.success("Perubahan halaman ini disimpan.")
                st.rerun()
            except Exception as e:
                st.error(f"Gagal menyimpan: {e}")

    with b2:
        if st.button("Hapus Terpilih", use_container_width=True, key="dosen_del_sel"):
            selected_ids = edited.loc[edited["Pilih"] == True, "ID Peg"].astype(str).tolist() if "ID Peg" in edited.columns else []
            if not selected_ids:
                st.warning("Tidak ada baris yang dipilih.")
            else:
                st.session_state.df_dosen = st.session_state.df_dosen[
                    ~st.session_state.df_dosen["ID Peg"].astype(str).isin(selected_ids)
                ].reset_index(drop=True)
                st.success(f"{len(selected_ids)} data dihapus.")
                st.rerun()

    with b3:
        if st.button("Delete All", use_container_width=True, key="dosen_del_all", type="primary"):
            st.session_state["_confirm_del_all_dosen"] = True

    with b4:
        ts = datetime.now().strftime("%Y%m%d_%H%M")
        export_df = ensure_pendidikan(st.session_state.df_dosen)
        cols = [c for c in display_cols if c in export_df.columns]
        excel_bytes = to_excel_bytes(export_df[cols] if cols else export_df, "Data Dosen")
        st.download_button(
            "Export Excel",
            data=excel_bytes,
            file_name=f"Data_Dosen_Periode_BKD_{ts}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    if st.session_state.get("_confirm_del_all_dosen"):
        st.warning("Yakin hapus **SEMUA** data Dosen?")
        c_yes, c_no = st.columns(2)
        with c_yes:
            if st.button("Ya, hapus semua Dosen", key="dosen_del_all_yes"):
                cols = list(st.session_state.df_dosen.columns)
                st.session_state.df_dosen = pd.DataFrame(columns=cols)
                st.session_state._confirm_del_all_dosen = False
                st.success("Semua data Dosen dihapus.")
                st.rerun()
        with c_no:
            if st.button("Batal", key="dosen_del_all_no"):
                st.session_state._confirm_del_all_dosen = False
                st.rerun()

    st.markdown("---")
    st.subheader("Ringkasan Dosen")
    c1, c2, c3, c4 = st.columns(4)
    dd = st.session_state.df_dosen
    with c1:
        st.metric("Total Dosen", len(dd))
    with c2:
        if "Status SERDOS" in dd.columns:
            n_ser = int(dd["Status SERDOS"].astype(str).str.lower().str.contains("sudah").sum())
            st.metric("Sudah SERDOS", n_ser)
        else:
            st.metric("SERDOS", "-")
    with c3:
        if "Jabatan Akademik" in dd.columns:
            st.metric(
                "Lektor / di atas",
                int(dd["Jabatan Akademik"].astype(str).str.lower().str.contains("lektor|guru besar|profesor").sum()),
            )
        else:
            st.metric("Jabatan", "-")
    with c4:
        st.metric("Eligible (estimasi)", len(get_eligible(dd)))

# ==================== KEKURANGAN ====================
elif menu == "Kekurangan Asesor BKD":
    render_header("Kekurangan Asesor BKD", "Ideal: setiap Prodi memiliki Asesor sebanyak 30% dari jumlah Dosen")

    st.info("💡 Idealnya, setiap Prodi memiliki Asesor sebanyak **30%** dari jumlah Dosen pada Prodi tersebut.")

    df_kekurangan = calculate_kekurangan(df_asesor, df_dosen)

    m1, m2, m3 = st.columns(3)
    m1.metric("Total Kekurangan Asesor", int(df_kekurangan["Kekurangan"].sum()))
    m2.metric("Prodi yang Kurang", len(df_kekurangan[df_kekurangan["Kekurangan"] > 0]))
    m3.metric("Prodi yang Cukup", len(df_kekurangan[df_kekurangan["Kekurangan"] == 0]))

    status_f = st.selectbox("Filter Status", ["Semua", "✅ Cukup", "⚠️ Kurang"])
    df_show = df_kekurangan if status_f == "Semua" else df_kekurangan[df_kekurangan["Status"] == status_f]

    st.dataframe(df_show, use_container_width=True, hide_index=True)

    ts = datetime.now().strftime("%Y%m%d_%H%M")
    excel_bytes = to_excel_bytes(df_kekurangan, "Kekurangan Asesor")
    st.download_button(
        "⬇️ Export Excel Kekurangan Asesor",
        data=excel_bytes,
        file_name=f"Kekurangan_Asesor_BKD_{ts}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.subheader("Perbandingan Jumlah Asesor vs Ideal (per Fakultas–Prodi)")
    df_k_plot = df_kekurangan.copy()
    df_k_plot["Label"] = df_k_plot["Fakultas"].astype(str) + " – " + df_k_plot["Prodi"].astype(str)
    fig = go.Figure()
    fig.add_trace(go.Bar(name="Jumlah Asesor", x=df_k_plot["Label"],
                         y=df_k_plot["Jumlah Asesor"], marker_color=PRIMARY))
    fig.add_trace(go.Bar(name="Jumlah Dosen", x=df_k_plot["Label"],
                         y=df_k_plot["Jumlah Dosen"], marker_color="#90caf9"))
    fig.add_trace(go.Scatter(name="Ideal 30%", x=df_k_plot["Label"],
                             y=df_k_plot["Ideal (30%)"], mode="lines+markers",
                             line=dict(color="#f5a623", width=3)))
    fig.update_layout(height=450, xaxis_tickangle=-45, barmode="group",
                      plot_bgcolor="white", paper_bgcolor="white",
                      legend=dict(orientation="h", y=1.08))
    st.plotly_chart(fig, use_container_width=True)

    kurang = df_kekurangan[df_kekurangan["Kekurangan"] > 0]
    if not kurang.empty:
        st.subheader("Detail Prodi yang Mengalami Kekurangan")
        for _, row in kurang.iterrows():
            with st.expander(f"{row['Prodi']} ({row['Fakultas']}) — Kurang {int(row['Kekurangan'])} Asesor"):
                st.write(f"- Jumlah Dosen: **{row['Jumlah Dosen']}**")
                st.write(f"- Jumlah Asesor saat ini: **{row['Jumlah Asesor']}**")
                st.write(f"- Ideal (30%): **{row['Ideal (30%)']}**")
                st.write(f"- Persentase Asesor: **{row['Persentase Asesor']}%**")
                st.progress(min(float(row["Persentase Asesor"]) / 30.0, 1.0))

# ==================== ELIGIBLE ====================
elif menu == "Eligible Asesor BKD":
    render_header("Eligible Asesor BKD", "Dosen yang memenuhi syarat menjadi Asesor BKD")

    st.markdown("""
    **Kriteria Eligible Asesor BKD:**
    1. **Status Dosen**: Dosen Tetap (aktif & tercatat di PDDIKTI)
    2. **Kualifikasi**: S3 + minimal **Lektor** · atau · S2 + minimal **Lektor Kepala**
    3. **Serdos**: Memiliki Sertifikat Pendidik (Status SERDOS = Sudah)
    """)

    df_eligible = get_eligible(df_dosen)

    e1, e2, e3 = st.columns(3)
    e1.metric("Total Eligible", len(df_eligible))
    e2.metric("% dari Total Dosen", f"{len(df_eligible)/max(len(df_dosen),1)*100:.1f}%")
    e3.metric("Total Dosen Periode", len(df_dosen))

    # Breakdown kenapa eligible bisa 0 (transparansi filter)
    with st.expander("ℹ️ Rincian filter kelayakan (diagnosa)"):
        dtmp = ensure_pendidikan(df_dosen)
        n_all = len(dtmp)
        n_tetap = int(dtmp.apply(lambda r: is_dosen_tetap(r.get("Status Dosen", "")), axis=1).sum()) if n_all else 0
        n_ser = int(dtmp.apply(lambda r: is_serdos_sudah(r.get("Status SERDOS", "")), axis=1).sum()) if n_all else 0
        n_both = int(dtmp.apply(
            lambda r: is_dosen_tetap(r.get("Status Dosen", "")) and is_serdos_sudah(r.get("Status SERDOS", "")),
            axis=1,
        ).sum()) if n_all else 0
        st.write(f"- Total dosen: **{n_all}**")
        st.write(f"- Status Dosen tetap/aktif: **{n_tetap}**")
        st.write(f"- Status SERDOS sudah: **{n_ser}**")
        st.write(f"- Tetap + Serdos: **{n_both}**")
        st.write(f"- Lolos kualifikasi jabatan + pendidikan: **{len(df_eligible)}**")
        if n_all and "Status Dosen" in dtmp.columns:
            st.caption("Nilai unik Status Dosen: " + ", ".join(sorted(dtmp["Status Dosen"].astype(str).unique())[:12]))
        if n_all and "Status SERDOS" in dtmp.columns:
            st.caption("Nilai unik Status SERDOS: " + ", ".join(sorted(dtmp["Status SERDOS"].astype(str).unique())[:12]))
        if n_all and "Jabatan Akademik" in dtmp.columns:
            st.caption("Nilai unik Jabatan: " + ", ".join(sorted(dtmp["Jabatan Akademik"].astype(str).unique())[:12]))

    ef1, ef2, ef3 = st.columns(3)
    with ef1:
        fak_f = st.multiselect("Fakultas", sorted(df_eligible["Fakultas"].unique()) if not df_eligible.empty else [])
    with ef2:
        pro_f = st.multiselect("Prodi", sorted(df_eligible["Prodi"].unique()) if not df_eligible.empty else [])
    with ef3:
        jab_f = st.multiselect("Jabatan", sorted(df_eligible["Jabatan Akademik"].unique()) if not df_eligible.empty else [])

    filtered_el = df_eligible.copy()
    if fak_f:
        filtered_el = filtered_el[filtered_el["Fakultas"].isin(fak_f)]
    if pro_f:
        filtered_el = filtered_el[filtered_el["Prodi"].isin(pro_f)]
    if jab_f:
        filtered_el = filtered_el[filtered_el["Jabatan Akademik"].isin(jab_f)]

    display_cols = [c for c in [
        "ID Peg", "NUPTK", "Nama", "Fakultas", "Prodi",
        "Jabatan Akademik", "Status Dosen", "Status SERDOS", "Pendidikan Terakhir"
    ] if c in filtered_el.columns]

    st.dataframe(filtered_el[display_cols], use_container_width=True, hide_index=True)

    ts = datetime.now().strftime("%Y%m%d_%H%M")
    excel_bytes = to_excel_bytes(filtered_el[display_cols] if not filtered_el.empty else filtered_el, "Eligible Asesor")
    st.download_button(
        "⬇️ Export Excel Eligible Asesor",
        data=excel_bytes,
        file_name=f"Eligible_Asesor_BKD_{ts}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    if not df_eligible.empty:
        st.subheader("Visualisasi Eligible")
        v1, v2 = st.columns(2)
        with v1:
            fig = px.bar(
                df_eligible.groupby("Fakultas").size().reset_index(name="Jumlah"),
                x="Fakultas", y="Jumlah", color="Jumlah",
                color_continuous_scale=["#a8d5e5", PRIMARY], title="Eligible per Fakultas"
            )
            st.plotly_chart(fig, use_container_width=True)
        with v2:
            fig = px.pie(df_eligible, names="Jabatan Akademik", title="Distribusi Jabatan Eligible",
                         hole=0.35, color_discrete_sequence=px.colors.sequential.Blues_r)
            st.plotly_chart(fig, use_container_width=True)

        already = set(df_asesor["ID Peg"].astype(str))
        df_eligible = df_eligible.copy()
        df_eligible["Sudah Asesor"] = df_eligible["ID Peg"].astype(str).isin(already)
        potensi = df_eligible[~df_eligible["Sudah Asesor"]]
        st.subheader("Potensi Rekrutmen")
        st.write(f"Dari **{len(df_eligible)}** eligible, **{len(potensi)}** belum menjadi Asesor.")
        if not potensi.empty:
            st.dataframe(potensi[display_cols], use_container_width=True, hide_index=True)

# ==================== FOOTER ====================
st.sidebar.markdown("---")
st.sidebar.caption("Dashboard Monitoring Asesor BKD")
st.sidebar.caption("Universitas Sebelas Maret (UNS)")
st.sidebar.caption("Tema warna: PEAR UNS (#0e7da7)")
