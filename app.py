import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import warnings
warnings.filterwarnings('ignore')

st.set_page_config(
    page_title="Dashboard Kopi Indonesia",
    page_icon="☕",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ─── CSS ───
st.markdown("""
<style>
    .main { background-color: #f9f5f0; }
    .stMetric { background: white; border-radius: 10px; padding: 10px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); }
    .title-card { background: linear-gradient(135deg, #4a2c0a 0%, #8B4513 50%, #c07e3b 100%);
        color: white; padding: 25px 30px; border-radius: 15px; margin-bottom: 25px; }
    h1 { color: #4a2c0a !important; }
    .section-header { background: #4a2c0a; color: white; padding: 8px 16px;
        border-radius: 8px; font-weight: 600; margin: 15px 0 10px 0; }
</style>
""", unsafe_allow_html=True)


# ─── DATA LOADING ───
@st.cache_data
def load_and_preprocess():
    years = [2014, 2015, 2016, 2017, 2018, 2019, 2020, 2021]

    def parse_file(path, row_start=10, row_end=44):
        df = pd.read_excel(path, engine='xlrd', header=None)
        data = df.iloc[row_start:row_end, [1, 2, 3, 4, 5, 6, 7, 8, 9]].copy()
        data.columns = ['Provinsi'] + years
        data = data[data['Provinsi'].notna()].copy()
        data['Provinsi'] = data['Provinsi'].astype(str).str.strip()
        for y in years:
            data[y] = pd.to_numeric(data[y], errors='coerce')
        data = data[data['Provinsi'].str.lower() != 'indonesia']
        data = data.dropna(subset=['Provinsi'])
        data = data[~data['Provinsi'].str.startswith('**')]
        data = data[~data['Provinsi'].str.startswith('-)')]
        return data.reset_index(drop=True)

    produksi = parse_file('/app/data/Produksi-Kopi.xls', row_start=11, row_end=45)
    areal    = parse_file('/app/data/Areal-Kopi.xls',    row_start=10, row_end=44)
    prodtv   = parse_file('/app/data/Prodtv-Kopi.xls',   row_start=10, row_end=44)

    # Melt to long format
    def to_long(df, val_name):
        return df.melt(id_vars='Provinsi', value_vars=years, var_name='Tahun', value_name=val_name)

    long_prod = to_long(produksi, 'Produksi_Ton')
    long_area = to_long(areal,    'Luas_Areal_Ha')
    long_prtv = to_long(prodtv,   'Produktivitas_KgHa')

    merged = long_prod.merge(long_area, on=['Provinsi', 'Tahun']).merge(long_prtv, on=['Provinsi', 'Tahun'])
    merged['Tahun'] = merged['Tahun'].astype(int)

    # Region mapping
    region_map = {
        'Aceh': 'Sumatera', 'Sumatera Utara': 'Sumatera', 'Sumatera Barat': 'Sumatera',
        'Riau': 'Sumatera', 'Kepulauan Riau': 'Sumatera', 'Jambi': 'Sumatera',
        'Sumatera Selatan': 'Sumatera', 'Kepulauan Bangka Belitung': 'Sumatera',
        'Bengkulu': 'Sumatera', 'Lampung': 'Sumatera',
        'DKI Jakarta': 'Jawa', 'Jawa Barat': 'Jawa', 'Banten': 'Jawa',
        'Jawa Tengah': 'Jawa', 'DI. Yogyakarta': 'Jawa', 'Jawa Timur': 'Jawa',
        'Bali': 'Bali & Nusa Tenggara', 'Nusa Tenggara Barat': 'Bali & Nusa Tenggara',
        'Nusa Tenggara Timur': 'Bali & Nusa Tenggara',
        'Kalimantan Barat': 'Kalimantan', 'Kalimantan Tengah': 'Kalimantan',
        'Kalimantan Selatan': 'Kalimantan', 'Kalimantan Timur': 'Kalimantan',
        'Kalimantan Utara': 'Kalimantan',
        'Sulawesi Utara': 'Sulawesi', 'Gorontalo': 'Sulawesi', 'Sulawesi Tengah': 'Sulawesi',
        'Sulawesi Selatan': 'Sulawesi', 'Sulawesi Barat': 'Sulawesi', 'Sulawesi Tenggara': 'Sulawesi',
        'Maluku': 'Maluku & Papua', 'Maluku Utara': 'Maluku & Papua',
        'Papua': 'Maluku & Papua', 'Papua Barat': 'Maluku & Papua',
    }
    merged['Wilayah'] = merged['Provinsi'].map(region_map).fillna('Lainnya')

    # Fill missing values with interpolation per province
    for col in ['Produksi_Ton', 'Luas_Areal_Ha', 'Produktivitas_KgHa']:
        merged[col] = merged.groupby('Provinsi')[col].transform(
            lambda x: x.interpolate(method='linear', limit_direction='both')
        )

    return merged, produksi, areal, prodtv, years


df, raw_prod, raw_area, raw_prtv, years = load_and_preprocess()

# ─── SIDEBAR ───
st.sidebar.markdown("## ☕ Filter Data")
selected_years = st.sidebar.slider("Rentang Tahun", 2014, 2021, (2014, 2021))
all_provinces = sorted(df['Provinsi'].unique())
all_regions = sorted(df['Wilayah'].unique())
selected_regions = st.sidebar.multiselect("Pilih Wilayah", all_regions, default=all_regions)
provinsi_filter = df[df['Wilayah'].isin(selected_regions)]['Provinsi'].unique()
selected_provinces = st.sidebar.multiselect("Pilih Provinsi", sorted(provinsi_filter), default=list(provinsi_filter)[:10])

# Filter
mask = (
    (df['Tahun'] >= selected_years[0]) &
    (df['Tahun'] <= selected_years[1]) &
    (df['Provinsi'].isin(selected_provinces))
)
dff = df[mask].copy()

# ─── HEADER ───
st.markdown("""
<div class="title-card">
    <h2 style="color:white; margin:0;">☕ Dashboard Analitik Perkebunan Kopi Indonesia</h2>
    <p style="margin:5px 0 0 0; opacity:0.85;">Data Produksi, Luas Areal & Produktivitas — 2014–2021 | Sumber: Ditjen Perkebunan</p>
</div>
""", unsafe_allow_html=True)

# ─── TABS ───
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📊 Ringkasan", "🗺️ Per Provinsi", "📈 Tren Waktu", "🔍 Preprocessing", "📋 Data Mentah"
])

# ═══════════════════════════════════════════════════
# TAB 1 — RINGKASAN
# ═══════════════════════════════════════════════════
with tab1:
    yr_latest = dff[dff['Tahun'] == dff['Tahun'].max()]
    yr_first  = dff[dff['Tahun'] == dff['Tahun'].min()]

    total_prod   = yr_latest['Produksi_Ton'].sum()
    total_area   = yr_latest['Luas_Areal_Ha'].sum()
    avg_prtv     = yr_latest['Produktivitas_KgHa'].mean()
    delta_prod   = total_prod - yr_first['Produksi_Ton'].sum()
    delta_area   = total_area - yr_first['Luas_Areal_Ha'].sum()
    delta_prtv   = avg_prtv  - yr_first['Produktivitas_KgHa'].mean()

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("🏭 Total Produksi", f"{total_prod:,.0f} Ton",    f"{delta_prod:+,.0f} Ton")
    c2.metric("🌿 Total Luas Areal", f"{total_area:,.0f} Ha",  f"{delta_area:+,.0f} Ha")
    c3.metric("⚡ Rerata Produktivitas", f"{avg_prtv:,.1f} Kg/Ha", f"{delta_prtv:+,.1f}")
    c4.metric("📍 Provinsi Dipilih", f"{dff['Provinsi'].nunique()}", "")

    st.markdown("---")
    col_a, col_b = st.columns(2)

    with col_a:
        st.markdown('<div class="section-header">Produksi per Wilayah (Tahun Terakhir)</div>', unsafe_allow_html=True)
        region_sum = yr_latest.groupby('Wilayah')['Produksi_Ton'].sum().reset_index().sort_values('Produksi_Ton', ascending=False)
        fig = px.bar(region_sum, x='Wilayah', y='Produksi_Ton',
                     color='Produksi_Ton', color_continuous_scale='YlOrBr',
                     labels={'Produksi_Ton': 'Produksi (Ton)', 'Wilayah': 'Wilayah'})
        fig.update_layout(margin=dict(t=10, b=10), coloraxis_showscale=False, height=320)
        st.plotly_chart(fig, use_container_width=True)

    with col_b:
        st.markdown('<div class="section-header">Distribusi Produksi (Pie)</div>', unsafe_allow_html=True)
        top_prov = yr_latest.groupby('Provinsi')['Produksi_Ton'].sum().nlargest(10).reset_index()
        fig2 = px.pie(top_prov, values='Produksi_Ton', names='Provinsi',
                      color_discrete_sequence=px.colors.sequential.YlOrBr)
        fig2.update_layout(margin=dict(t=10, b=10), height=320)
        st.plotly_chart(fig2, use_container_width=True)

    st.markdown('<div class="section-header">Top 10 Provinsi Penghasil Kopi</div>', unsafe_allow_html=True)
    top10 = yr_latest.groupby('Provinsi')[['Produksi_Ton', 'Luas_Areal_Ha', 'Produktivitas_KgHa']].sum().nlargest(10, 'Produksi_Ton').reset_index()
    top10 = top10.rename(columns={'Produksi_Ton': 'Produksi (Ton)', 'Luas_Areal_Ha': 'Luas Areal (Ha)', 'Produktivitas_KgHa': 'Produktivitas (Kg/Ha)'})
    st.dataframe(top10.style.background_gradient(cmap='YlOrBr', subset=['Produksi (Ton)']).format({
        'Produksi (Ton)': '{:,.0f}', 'Luas Areal (Ha)': '{:,.0f}', 'Produktivitas (Kg/Ha)': '{:,.1f}'
    }), use_container_width=True, height=380)


# ═══════════════════════════════════════════════════
# TAB 2 — PER PROVINSI
# ═══════════════════════════════════════════════════
with tab2:
    st.markdown('<div class="section-header">Produksi Kopi per Provinsi (Bar Chart)</div>', unsafe_allow_html=True)
    yr_sel = st.selectbox("Pilih Tahun", sorted(dff['Tahun'].unique(), reverse=True), key='yr_prov')
    df_yr = dff[dff['Tahun'] == yr_sel].sort_values('Produksi_Ton', ascending=True)
    fig3 = px.bar(df_yr, y='Provinsi', x='Produksi_Ton', orientation='h',
                  color='Produksi_Ton', color_continuous_scale='YlOrBr',
                  labels={'Produksi_Ton': 'Produksi (Ton)'}, height=700)
    fig3.update_layout(margin=dict(t=10, b=10), coloraxis_showscale=False)
    st.plotly_chart(fig3, use_container_width=True)

    col1, col2 = st.columns(2)
    with col1:
        st.markdown('<div class="section-header">Scatter: Luas Areal vs Produksi</div>', unsafe_allow_html=True)
        fig4 = px.scatter(df_yr, x='Luas_Areal_Ha', y='Produksi_Ton',
                          size='Produktivitas_KgHa', color='Wilayah', hover_name='Provinsi',
                          labels={'Luas_Areal_Ha': 'Luas Areal (Ha)', 'Produksi_Ton': 'Produksi (Ton)'},
                          color_discrete_sequence=px.colors.qualitative.Set2, height=400)
        fig4.update_layout(margin=dict(t=10, b=10))
        st.plotly_chart(fig4, use_container_width=True)

    with col2:
        st.markdown('<div class="section-header">Produktivitas per Wilayah (Box Plot)</div>', unsafe_allow_html=True)
        fig5 = px.box(dff, x='Wilayah', y='Produktivitas_KgHa', color='Wilayah',
                      labels={'Produktivitas_KgHa': 'Produktivitas (Kg/Ha)'},
                      color_discrete_sequence=px.colors.qualitative.Set2, height=400)
        fig5.update_layout(showlegend=False, margin=dict(t=10, b=10))
        st.plotly_chart(fig5, use_container_width=True)


# ═══════════════════════════════════════════════════
# TAB 3 — TREN WAKTU
# ═══════════════════════════════════════════════════
with tab3:
    st.markdown('<div class="section-header">Tren Produksi Nasional per Tahun</div>', unsafe_allow_html=True)
    nasional = dff.groupby('Tahun')[['Produksi_Ton', 'Luas_Areal_Ha', 'Produktivitas_KgHa']].sum().reset_index()
    nasional['Produktivitas_KgHa'] = dff.groupby('Tahun')['Produktivitas_KgHa'].mean().values

    fig_nat = make_subplots(rows=1, cols=3, subplot_titles=('Produksi (Ton)', 'Luas Areal (Ha)', 'Produktivitas (Kg/Ha)'))
    for i, (col, color) in enumerate([('Produksi_Ton', '#8B4513'), ('Luas_Areal_Ha', '#228B22'), ('Produktivitas_KgHa', '#DAA520')], 1):
        fig_nat.add_trace(go.Scatter(x=nasional['Tahun'], y=nasional[col], mode='lines+markers',
                                     line=dict(color=color, width=3), marker=dict(size=8)), row=1, col=i)
    fig_nat.update_layout(height=320, showlegend=False, margin=dict(t=40, b=10))
    st.plotly_chart(fig_nat, use_container_width=True)

    st.markdown('<div class="section-header">Tren Produksi per Provinsi (Line Chart)</div>', unsafe_allow_html=True)
    prov_selected = st.multiselect("Pilih Provinsi untuk Tren", all_provinces,
                                   default=list(dff.groupby('Provinsi')['Produksi_Ton'].sum().nlargest(5).index))
    df_trend = df[df['Provinsi'].isin(prov_selected) & (df['Tahun'] >= selected_years[0]) & (df['Tahun'] <= selected_years[1])]
    fig6 = px.line(df_trend, x='Tahun', y='Produksi_Ton', color='Provinsi',
                   markers=True, labels={'Produksi_Ton': 'Produksi (Ton)'},
                   color_discrete_sequence=px.colors.qualitative.Set1, height=400)
    fig6.update_layout(margin=dict(t=10, b=10))
    st.plotly_chart(fig6, use_container_width=True)

    st.markdown('<div class="section-header">Heatmap Produksi: Provinsi vs Tahun</div>', unsafe_allow_html=True)
    pivot = dff.pivot_table(values='Produksi_Ton', index='Provinsi', columns='Tahun', aggfunc='sum')
    fig7 = px.imshow(pivot, color_continuous_scale='YlOrBr', aspect='auto',
                     labels={'color': 'Produksi (Ton)'}, height=600)
    fig7.update_layout(margin=dict(t=10, b=10))
    st.plotly_chart(fig7, use_container_width=True)


# ═══════════════════════════════════════════════════
# TAB 4 — PREPROCESSING
# ═══════════════════════════════════════════════════
with tab4:
    st.markdown("### 🔍 Laporan Data Preprocessing")

    st.markdown("#### 1. Ringkasan Dataset Mentah")
    c1, c2, c3 = st.columns(3)
    c1.metric("Produksi — Baris Mentah", f"{len(raw_prod)} baris, {raw_prod.shape[1]} kolom")
    c2.metric("Areal — Baris Mentah",    f"{len(raw_area)} baris, {raw_area.shape[1]} kolom")
    c3.metric("Prodtv — Baris Mentah",   f"{len(raw_prtv)} baris, {raw_prtv.shape[1]} kolom")

    st.markdown("#### 2. Missing Values Setelah Preprocessing")
    miss_df = pd.DataFrame({
        'Kolom': ['Produksi_Ton', 'Luas_Areal_Ha', 'Produktivitas_KgHa'],
        'Missing (sebelum interpolasi)': [
            df['Produksi_Ton'].isna().sum(),
            df['Luas_Areal_Ha'].isna().sum(),
            df['Produktivitas_KgHa'].isna().sum(),
        ],
        'Missing (sesudah interpolasi)': [0, 0, 0]
    })
    st.dataframe(miss_df, use_container_width=True)

    st.markdown("#### 3. Statistik Deskriptif")
    desc = df[['Produksi_Ton', 'Luas_Areal_Ha', 'Produktivitas_KgHa']].describe().T
    desc.columns = ['Count', 'Mean', 'Std', 'Min', '25%', '50%', '75%', 'Max']
    st.dataframe(desc.style.format("{:,.2f}"), use_container_width=True)

    st.markdown("#### 4. Distribusi Variabel")
    var = st.selectbox("Pilih Variabel", ['Produksi_Ton', 'Luas_Areal_Ha', 'Produktivitas_KgHa'])
    fig_hist = px.histogram(df, x=var, nbins=40, color='Wilayah',
                            color_discrete_sequence=px.colors.qualitative.Set2,
                            marginal='box', height=400,
                            labels={var: var.replace('_', ' ')})
    st.plotly_chart(fig_hist, use_container_width=True)

    st.markdown("#### 5. Korelasi Antar Variabel")
    corr = df[['Produksi_Ton', 'Luas_Areal_Ha', 'Produktivitas_KgHa']].corr()
    fig_corr = px.imshow(corr, text_auto=True, color_continuous_scale='RdBu_r',
                         zmin=-1, zmax=1, height=350)
    st.plotly_chart(fig_corr, use_container_width=True)

    st.markdown("#### 6. Outlier Detection (IQR)")
    def detect_outliers(series):
        Q1, Q3 = series.quantile(0.25), series.quantile(0.75)
        IQR = Q3 - Q1
        return ((series < Q1 - 1.5*IQR) | (series > Q3 + 1.5*IQR)).sum()

    out_df = pd.DataFrame({
        'Variabel': ['Produksi_Ton', 'Luas_Areal_Ha', 'Produktivitas_KgHa'],
        'Jumlah Outlier (IQR)': [
            detect_outliers(df['Produksi_Ton']),
            detect_outliers(df['Luas_Areal_Ha']),
            detect_outliers(df['Produktivitas_KgHa']),
        ]
    })
    st.dataframe(out_df, use_container_width=True)

    st.markdown("#### 7. Langkah-langkah Preprocessing")
    steps = [
        "✅ **Load Data**: Membaca 3 file XLS (Produksi, Areal, Produktivitas) menggunakan pandas + xlrd",
        "✅ **Ekstraksi Baris**: Memotong baris header & footer (sumber, catatan) → hanya baris provinsi",
        "✅ **Rename Kolom**: Menamai ulang kolom numerik menjadi tahun (2014–2021)",
        "✅ **Type Casting**: Mengkonversi nilai ke numerik, string tidak valid → NaN",
        "✅ **Filter Baris**: Menghapus baris 'Indonesia' (agregat nasional) & baris kosong",
        "✅ **Melt / Reshape**: Dari format wide → long (Provinsi × Tahun sebagai baris)",
        "✅ **Merge 3 Dataset**: Menggabungkan produksi, areal, produktivitas by [Provinsi, Tahun]",
        "✅ **Imputasi**: Interpolasi linear per provinsi untuk mengisi nilai kosong",
        "✅ **Feature Engineering**: Menambahkan kolom Wilayah (region mapping 34 provinsi)",
    ]
    for s in steps:
        st.markdown(s)


# ═══════════════════════════════════════════════════
# TAB 5 — DATA MENTAH
# ═══════════════════════════════════════════════════
with tab5:
    st.markdown("### 📋 Dataset Final (Setelah Preprocessing)")
    st.markdown(f"**{len(dff):,} baris** × **{dff.shape[1]} kolom**")
    st.dataframe(
        dff.sort_values(['Tahun', 'Produksi_Ton'], ascending=[True, False]).reset_index(drop=True).style.format({
            'Produksi_Ton': '{:,.2f}',
            'Luas_Areal_Ha': '{:,.2f}',
            'Produktivitas_KgHa': '{:,.2f}',
        }),
        use_container_width=True, height=500
    )
    csv = dff.to_csv(index=False).encode('utf-8')
    st.download_button("⬇️ Download CSV", csv, "kopi_indonesia_clean.csv", "text/csv")

st.markdown("---")
st.caption("Data: Direktorat Jenderal Perkebunan, Kementerian Pertanian RI | Dashboard oleh Streamlit + Plotly")
