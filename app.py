import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path
import io

# --- KONFIGURASI HALAMAN ---
st.set_page_config(
    page_title="Dashboard Monitoring Kegiatan & Anggaran",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- CUSTOM CSS MODERN (DARK & LIGHT THEME FRIENDLY, ZERO FLICKER) ---
st.markdown("""
<style>
    .kpi-card {
        background-color: var(--background-color, #ffffff);
        border-radius: 12px;
        padding: 18px 22px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.08), 0 2px 4px -1px rgba(0, 0, 0, 0.04);
        border: 1px solid rgba(148, 163, 184, 0.2);
        transition: transform 0.2s ease, box-shadow 0.2s ease;
        margin-bottom: 12px;
    }
    .kpi-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.12);
    }
    .kpi-title {
        font-size: 0.82rem;
        font-weight: 700;
        color: #94a3b8;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        margin-bottom: 6px;
    }
    .kpi-value {
        font-size: 1.85rem;
        font-weight: 800;
        line-height: 1.2;
    }
    .kpi-subtext {
        font-size: 0.88rem;
        margin-top: 6px;
        color: #cbd5e1;
    }
    .badge {
        display: inline-block;
        padding: 2px 8px;
        border-radius: 6px;
        font-weight: 600;
        font-size: 0.8rem;
    }
    .badge-success {
        background-color: rgba(34, 197, 94, 0.15);
        color: #22c55e;
        border: 1px solid rgba(34, 197, 94, 0.3);
    }
    .badge-warning {
        background-color: rgba(245, 158, 11, 0.15);
        color: #f59e0b;
        border: 1px solid rgba(245, 158, 11, 0.3);
    }
    .badge-danger {
        background-color: rgba(239, 68, 68, 0.15);
        color: #ef4444;
        border: 1px solid rgba(239, 68, 68, 0.3);
    }
    .trace-card {
        background-color: rgba(59, 130, 246, 0.08);
        border: 1px solid rgba(59, 130, 246, 0.25);
        border-radius: 8px;
        padding: 10px 14px;
        font-size: 0.85rem;
        margin-bottom: 12px;
    }
</style>
""", unsafe_allow_html=True)


# --- FUNGSI PARSING & PEMBERSIHAN DATA CERDAS ---
@st.cache_data(show_spinner=False)
def read_excel_file(file_bytes):
    """Membaca seluruh daftar sheet dari file Excel"""
    excel_data = pd.ExcelFile(io.BytesIO(file_bytes))
    return excel_data.sheet_names


@st.cache_data(show_spinner=False)
def smart_load_sheet(file_bytes, sheet_name):
    """
    Mendeteksi posisi header baris data secara otomatis (Header Sniffer)
    sehingga sheet pivot atau sheet dengan judul di baris atas tetap terbaca rapi.
    """
    try:
        df_sample = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, nrows=10, header=None)
        if df_sample.empty:
            return pd.DataFrame()
        
        # Cari baris yang paling banyak memiliki teks header non-kosong
        best_header_row = 0
        max_non_null = 0
        for idx, row in df_sample.iterrows():
            non_null_count = sum(1 for val in row if pd.notna(val) and str(val).strip() != "" and not str(val).startswith("Unnamed:"))
            if non_null_count > max_non_null:
                max_non_null = non_null_count
                best_header_row = idx
                
        df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=best_header_row)
        df.columns = [str(c).strip() for c in df.columns]
        return df
    except Exception as e:
        st.error(f"Gagal membaca sheet '{sheet_name}': {e}")
        return pd.DataFrame()


def clean_number(value):
    """Membersihkan format angka string/ribuan/persen ke float murni"""
    if pd.isna(value):
        return 0.0
    if isinstance(value, (int, float, np.integer, np.floating)):
        return float(value)
    value = str(value).strip().replace("%", "").replace(" ", "")
    if value in ["", "-", "–", "—", "nan", "None"]:
        return 0.0
    if "," in value and "." in value:
        if value.rfind(",") > value.rfind("."):
            value = value.replace(".", "").replace(",", ".")
        else:
            value = value.replace(",", "")
    elif "," in value:
        parts = value.split(",")
        if len(parts[-1]) <= 2:
            value = value.replace(",", ".")
        else:
            value = value.replace(",", "")
    elif value.count(".") > 1:
        value = value.replace(".", "")
    try:
        return float(value)
    except:
        return 0.0


def auto_detect_columns(cols):
    """
    SMART TRACER: Melacak dan mencocokkan nama kolom secara otomatis
    berdasarkan variasi sinonim (Indonesia & Inggris) tanpa perlu setting manual.
    """
    mapping = {
        "ssr": None, "bulan": None, "bl": None,
        "jml_kegiatan_planning": None, "unit_cost": None,
        "total_budget": None, "jml_kegiatan_realisasi": None,
        "realisasi": None, "justifikasi": None
    }
    
    patterns = {
        "ssr": ["ssr", "sub recipient", "sub-recipient", "lembaga", "organisasi", "mitra", "unit", "cabang", "wilayah", "agency"],
        "bulan": ["bulan", "month", "periode", "waktu", "bln", "date", "tanggal"],
        "bl": ["bl", "budget line", "kode bl", "pos anggaran", "kategori bl", "activity code", "mata anggaran"],
        "jml_kegiatan_planning": ["jml kegiatan p", "planning", "target kegiatan", "rencana", "target", "plan qty", "volume plan", "jml p", "plan"],
        "unit_cost": ["unit cost", "biaya satuan", "harga satuan", "tarif", "unit_cost", "cost"],
        "total_budget": ["total budget", "total dana planning", "total budget planning", "pagu", "total anggaran", "budget plan", "total"],
        "jml_kegiatan_realisasi": ["jml kegiatan r", "realisasi kegiatan", "capaian kegiatan", "actual qty", "volume real", "jml r", "terlaksana"],
        "realisasi": ["realisasi", "serapan", "actual budget", "pengeluaran", "spending", "amount", "dana realisasi", "total realisasi", "actual cost"],
        "justifikasi": ["justifikasi", "alasan", "keterangan", "catatan", "remark", "notes", "reason", "penjelasan", "deskripsi"]
    }
    
    clean_cols = [str(c).strip() for c in cols if pd.notna(c) and not str(c).startswith("Unnamed:")]
    
    for key, kw_list in patterns.items():
        for c in clean_cols:
            low = c.lower()
            # Cek exact match atau keyword match
            if any(kw == low for kw in kw_list) or any(kw in low for kw in kw_list):
                mapping[key] = c
                break
                
    return mapping


def process_clean_dataframe(df_raw, col_map, month_range_option, apply_outlier_fix):
    """Pipeline pemrosesan dan standardisasi data otomatis"""
    df = df_raw.copy()

    # Rename sesuai mapping yang terdeteksi
    inv_map = {v: k for k, v in col_map.items() if v is not None and v in df.columns}
    df = df.rename(columns=inv_map)

    # Pastikan seluruh kolom standar ada di dataframe
    for required in ["ssr", "bulan", "bl", "jml_kegiatan_planning", "unit_cost", "total_budget", "jml_kegiatan_realisasi", "realisasi"]:
        if required not in df.columns:
            df[required] = 0.0 if required in ["bl", "jml_kegiatan_planning", "unit_cost", "total_budget", "jml_kegiatan_realisasi", "realisasi"] else "-"

    # Treatment numerik
    num_cols = ["bl", "jml_kegiatan_planning", "unit_cost", "total_budget", "jml_kegiatan_realisasi", "realisasi"]
    for col in num_cols:
        df[col] = df[col].apply(clean_number)

    # Mapping Bulan
    month_map = {
        "jan": 1, "feb": 2, "mar": 3, "apr": 4, "mei": 5, "may": 5,
        "jun": 6, "jul": 7, "agu": 8, "ags": 8, "aug": 8, "sep": 9,
        "okt": 10, "oct": 10, "nov": 11, "des": 12, "dec": 12
    }
    def get_month_num(val):
        if pd.isna(val):
            return 1
        s = str(val).lower().replace(".", " ")
        for k, n in month_map.items():
            if k in s:
                return n
        try:
            val_int = int(val)
            if 1 <= val_int <= 12:
                return val_int
        except:
            pass
        return 1

    df["bulan_no"] = df["bulan"].apply(get_month_num)
    month_names = {
        1: "Januari", 2: "Februari", 3: "Maret", 4: "April",
        5: "Mei", 6: "Juni", 7: "Juli", 8: "Agustus",
        9: "September", 10: "Oktober", 11: "November", 12: "Desember"
    }
    df["bulan_nama"] = df["bulan_no"].map(month_names)

    # Filter Periode
    if month_range_option == "Semester 1 (Januari – Juni)":
        df = df[df["bulan_no"].between(1, 6)].copy()
    elif month_range_option == "Semester 2 (Juli – Desember)":
        df = df[df["bulan_no"].between(7, 12)].copy()

    # Treatment outlier
    if apply_outlier_fix:
        for idx in df[df["jml_kegiatan_planning"] > 1000].index:
            uc = df.loc[idx, "unit_cost"]
            tb = df.loc[idx, "total_budget"]
            p = df.loc[idx, "jml_kegiatan_planning"]
            r_keg = df.loc[idx, "jml_kegiatan_realisasi"]
            if uc > 0 and tb > 0:
                calc_p = round(tb / uc)
                if calc_p < 1000:
                    df.loc[idx, "jml_kegiatan_planning"] = calc_p
                elif p > 0 and round(tb / p) < 1000:
                    df.loc[idx, "jml_kegiatan_planning"] = round(tb / p)
                    df.loc[idx, "unit_cost"] = p
                else:
                    df.loc[idx, "jml_kegiatan_planning"] = max(1.0, r_keg)
            else:
                df.loc[idx, "jml_kegiatan_planning"] = max(1.0, r_keg)

    # Valid total budget
    df["total_budget_valid"] = np.where(df["total_budget"] > 0, df["total_budget"], df["jml_kegiatan_planning"] * df["unit_cost"])
    
    # Kolom teks
    df["ssr"] = df["ssr"].astype(str).str.strip()
    if "justifikasi" in df.columns:
        df["justifikasi"] = df["justifikasi"].fillna("-").astype(str).str.strip()
    else:
        df["justifikasi"] = "-"

    return df.reset_index(drop=True)


# ==============================================================================
# SIDEBAR: SUMBER DATA & SMART HEADER TRACER
# ==============================================================================
st.sidebar.markdown("### 📁 Sumber Data Excel")

# 1. Upload File Excel
uploaded_file = st.sidebar.file_uploader(
    "Unggah File Excel (Opsional):",
    type=["xlsx", "xls"],
    help="Pilih file Excel Anda atau gunakan data default"
)

file_bytes = None
file_source_name = "Dashboard baru.xlsx"

if uploaded_file is not None:
    file_bytes = uploaded_file.getvalue()
    file_source_name = uploaded_file.name
else:
    default_path = Path("Dashboard baru.xlsx")
    if not default_path.exists() and Path("Dashboard baru(1).xlsx").exists():
        default_path = Path("Dashboard baru(1).xlsx")
        file_source_name = "Dashboard baru(1).xlsx"
        
    if default_path.exists():
        with open(default_path, "rb") as f:
            file_bytes = f.read()
    else:
        st.error("⚠️ File Excel tidak ditemukan. Silakan unggah file Excel pada menu di sidebar.")
        st.stop()

# 2. Pemilihan Sheet Dinamis
available_sheets = read_excel_file(file_bytes)
default_sheet_index = available_sheets.index("Activity") if "Activity" in available_sheets else 0

selected_sheet = st.sidebar.selectbox(
    "📑 Pilih Sheet Data:",
    options=available_sheets,
    index=default_sheet_index,
    help="Pilih sheet yang ingin Anda analisa"
)

# Load data mentah dengan Smart Header Sniffer
df_raw = smart_load_sheet(file_bytes, selected_sheet)

# Otomatis deteksi & trace kolom sheet yang terpilih
auto_map = auto_detect_columns(df_raw.columns.tolist())
traced_count = sum(1 for v in auto_map.values() if v is not None)

# 3. Menu Modifikasi & Mapping Kolom Dinamis (Auto-Trace)
with st.sidebar.expander("⚙️ Modifikasi & Tracing Kolom", expanded=False):
    st.markdown(f"""
    <div class="trace-card">
        <b>🔍 Status Auto-Trace Sheet:</b><br>
        Sheet <code>{selected_sheet}</code> memiliki <b>{len(df_raw.columns)} kolom</b>.<br>
        Terdeteksi otomatis: <b>{traced_count} dari 9 kolom utama</b>.
    </div>
    """, unsafe_allow_html=True)

    st.markdown("**1. Periode Bulan Analisis:**")
    month_range_choice = st.selectbox(
        "Rentang Periode:",
        options=[
            "Semester 1 (Januari – Juni)",
            "Semua Bulan (Januari – Desember)",
            "Semester 2 (Juli – Desember)"
        ],
        index=0,
        key=f"period_{selected_sheet}"
    )

    st.markdown("**2. Pembersihan Nilai Outlier:**")
    apply_outlier = st.checkbox(
        "Normalisasi Target Tertukar (Outlier Fix)",
        value=True,
        key=f"outlier_{selected_sheet}"
    )

    st.markdown("**3. Mapping Kolom Otomatis (Bisa disesuaikan jika perlu):**")
    raw_columns = ["<Tidak Digunakan>"] + df_raw.columns.tolist()

    def get_index_for(col_key):
        val = auto_map.get(col_key)
        return raw_columns.index(val) if val in raw_columns else 0

    custom_col_map = {}
    col_labels = {
        "ssr": "Kolom Lembaga / SSR:",
        "bulan": "Kolom Bulan / Periode:",
        "bl": "Kolom Budget Line (BL):",
        "jml_kegiatan_planning": "Kolom Target Kegiatan (P):",
        "unit_cost": "Kolom Unit Cost:",
        "total_budget": "Kolom Total Budget:",
        "jml_kegiatan_realisasi": "Kolom Realisasi Kegiatan (R):",
        "realisasi": "Kolom Realisasi Dana / Amount:",
        "justifikasi": "Kolom Justifikasi / Keterangan:"
    }

    for col_key, label in col_labels.items():
        custom_col_map[col_key] = st.selectbox(
            label,
            options=raw_columns,
            index=get_index_for(col_key),
            key=f"col_{selected_sheet}_{col_key}"
        )

    active_col_map = {k: (v if v != "<Tidak Digunakan>" else None) for k, v in custom_col_map.items()}

# Jalankan pembersihan data
df_clean = process_clean_dataframe(df_raw, active_col_map, month_range_choice, apply_outlier)

st.sidebar.markdown("---")


# ==============================================================================
# FRAGMENT DASHBOARD (ZERO-RELOAD / DYNAMIC REACTION)
# ==============================================================================
@st.fragment
def render_interactive_dashboard(df_base):
    st.sidebar.markdown("### 🎛️ Filter Analisis")
    
    # 1. Filter Periode Bulan
    available_months = df_base.sort_values("bulan_no")["bulan_nama"].dropna().unique().tolist()
    selected_months = st.sidebar.multiselect(
        "Pilih Periode Bulan:",
        options=available_months,
        default=available_months,
        help="Pilih satu atau lebih bulan"
    )

    # 2. Filter SSR
    available_ssr = sorted(df_base["ssr"].dropna().unique().tolist())
    selected_ssr = st.sidebar.multiselect(
        "Pilih SSR (Sub-Sub Recipient):",
        options=available_ssr,
        default=available_ssr
    )

    # 3. Filter Budget Line (BL)
    available_bl = sorted(df_base["bl"].dropna().unique().astype(int).tolist())
    select_all_bl_toggle = st.sidebar.checkbox("Pilih Semua Budget Line (BL)", value=True)
    
    if select_all_bl_toggle:
        selected_bl = available_bl
    else:
        selected_bl = st.sidebar.multiselect(
            "Pilih BL Tertentu:",
            options=available_bl,
            default=available_bl[:10] if len(available_bl) >= 10 else available_bl
        )

    # Terapkan Filter
    df_filtered = df_base[
        (df_base["bulan_nama"].isin(selected_months)) &
        (df_base["ssr"].isin(selected_ssr)) &
        (df_base["bl"].isin(selected_bl))
    ]

    st.markdown(f"**Sumber Data:** `{file_source_name}` (Sheet: `{selected_sheet}`) | **Data Ditampilkan:** `{len(df_filtered):,}` baris")

    if df_filtered.empty:
        st.warning("⚠️ Tidak ada data yang sesuai dengan kombinasi filter yang dipilih.")
        return

    st.markdown("---")

    # ==========================================================================
    # 1. MAIN KPI: JUMLAH KEGIATAN TERLAKSANA & BUDGET BESERTA PERSENTASENYA
    # ==========================================================================
    st.subheader("🎯 Main KPI (Indikator Utama)")

    tot_keg_plan = df_filtered["jml_kegiatan_planning"].sum()
    tot_keg_real = df_filtered["jml_kegiatan_realisasi"].sum()
    persen_kegiatan = (tot_keg_real / tot_keg_plan * 100) if tot_keg_plan > 0 else 0

    tot_dana_plan = df_filtered["total_budget_valid"].sum()
    tot_dana_real = df_filtered["realisasi"].sum()
    persen_dana = (tot_dana_real / tot_dana_plan * 100) if tot_dana_plan > 0 else 0
    sisa_dana = tot_dana_plan - tot_dana_real

    col_kpi1, col_kpi2, col_kpi3 = st.columns(3)

    with col_kpi1:
        st.markdown(f"""
        <div class="kpi-card" style="border-top: 4px solid #3b82f6;">
            <div class="kpi-title">1. Jumlah Kegiatan Terlaksana</div>
            <div class="kpi-value" style="color: #3b82f6;">{persen_kegiatan:.1f}%</div>
            <div class="kpi-subtext">
                <b>{tot_keg_real:,.0f}</b> terlaksana dari <b>{tot_keg_plan:,.0f}</b> target kegiatan
            </div>
        </div>
        """, unsafe_allow_html=True)

    with col_kpi2:
        color_val = "#10b981" if persen_dana <= 100 else "#f59e0b"
        st.markdown(f"""
        <div class="kpi-card" style="border-top: 4px solid #10b981;">
            <div class="kpi-title">2. Budget & Serapan Dana</div>
            <div class="kpi-value" style="color: {color_val};">{persen_dana:.1f}%</div>
            <div class="kpi-subtext">
                <b>Rp {tot_dana_real:,.0f}</b> realisasi dari <b>Rp {tot_dana_plan:,.0f}</b> budget
            </div>
        </div>
        """, unsafe_allow_html=True)

    with col_kpi3:
        status_lbl = "Tersisa" if sisa_dana >= 0 else "Overbudget"
        badge_stat = "badge-success" if sisa_dana >= 0 else "badge-danger"
        st.markdown(f"""
        <div class="kpi-card" style="border-top: 4px solid #f59e0b;">
            <div class="kpi-title">3. Sisa / Deviasi Budget</div>
            <div class="kpi-value">Rp {abs(sisa_dana):,.0f}</div>
            <div class="kpi-subtext">
                Status: <span class="badge {badge_stat}">{status_lbl}</span>
            </div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ==========================================================================
    # 2. DETAIL BUDGET VS SERAPAN PER BULAN
    # ==========================================================================
    st.subheader("📅 Detail Budget vs Serapan per Bulan")
    
    df_monthly = df_filtered.groupby(["bulan_no", "bulan_nama"]).agg(
        tot_budget=("total_budget_valid", "sum"),
        tot_realisasi=("realisasi", "sum"),
        tot_plan_keg=("jml_kegiatan_planning", "sum"),
        tot_real_keg=("jml_kegiatan_realisasi", "sum")
    ).reset_index().sort_values("bulan_no")

    df_monthly["% Serapan Dana"] = np.where(df_monthly["tot_budget"] > 0, df_monthly["tot_realisasi"] / df_monthly["tot_budget"] * 100, 0)
    df_monthly["% Capaian Kegiatan"] = np.where(df_monthly["tot_plan_keg"] > 0, df_monthly["tot_real_keg"] / df_monthly["tot_plan_keg"] * 100, 0)

    col_m_chart, col_m_table = st.columns([6, 4])

    with col_m_chart:
        fig_monthly = go.Figure()
        fig_monthly.add_trace(go.Bar(
            x=df_monthly["bulan_nama"],
            y=df_monthly["tot_budget"],
            name="Budget Planning",
            marker_color="#94a3b8"
        ))
        fig_monthly.add_trace(go.Bar(
            x=df_monthly["bulan_nama"],
            y=df_monthly["tot_realisasi"],
            name="Realisasi Dana",
            marker_color="#10b981"
        ))
        fig_monthly.update_layout(
            barmode="group",
            title="Tren Budget vs Serapan Realisasi per Bulan (Rp)",
            xaxis_title="Bulan",
            yaxis_title="Nominal Dana (Rp)",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            height=360,
            margin=dict(l=20, r=20, t=40, b=20)
        )
        st.plotly_chart(fig_monthly, use_container_width=True)

    with col_m_table:
        st.markdown("**Tabel Rincian Serapan per Bulan:**")
        df_m_show = df_monthly[["bulan_nama", "tot_realisasi", "tot_budget", "% Serapan Dana"]].copy()
        df_m_show.columns = ["Bulan", "Realisasi (Rp)", "Budget (Rp)", "% Serapan"]
        st.dataframe(
            df_m_show.style.format({
                "Realisasi (Rp)": "Rp {:,.0f}",
                "Budget (Rp)": "Rp {:,.0f}",
                "% Serapan": "{:.2f}%"
            }),
            use_container_width=True,
            hide_index=True,
            height=320
        )

    st.markdown("<br>", unsafe_allow_html=True)

    # ==========================================================================
    # 3. BUDGET VS SERAPAN PER SSR
    # ==========================================================================
    st.subheader("🏢 Budget vs Serapan per SSR")

    tab_ssr_dana, tab_ssr_keg = st.tabs([
        "💰 Budget vs Serapan Dana per SSR",
        "📌 Kegiatan Terlaksana vs Target per SSR"
    ])

    df_ssr = df_filtered.groupby("ssr").agg(
        tot_plan_keg=("jml_kegiatan_planning", "sum"),
        tot_real_keg=("jml_kegiatan_realisasi", "sum"),
        tot_plan_dana=("total_budget_valid", "sum"),
        tot_real_dana=("realisasi", "sum")
    ).reset_index()

    df_ssr["% Serapan Dana"] = np.where(df_ssr["tot_plan_dana"] > 0, df_ssr["tot_real_dana"] / df_ssr["tot_plan_dana"] * 100, 0)
    df_ssr["% Serapan Kegiatan"] = np.where(df_ssr["tot_plan_keg"] > 0, df_ssr["tot_real_keg"] / df_ssr["tot_plan_keg"] * 100, 0)

    with tab_ssr_dana:
        col_c_dana, col_t_dana = st.columns([6, 4])
        with col_c_dana:
            fig_dana = go.Figure()
            fig_dana.add_trace(go.Bar(
                x=df_ssr["ssr"],
                y=df_ssr["tot_plan_dana"],
                name="Total Budget",
                marker_color="#cbd5e1"
            ))
            fig_dana.add_trace(go.Bar(
                x=df_ssr["ssr"],
                y=df_ssr["tot_real_dana"],
                name="Dana Realisasi",
                marker_color="#10b981"
            ))
            fig_dana.update_layout(
                barmode="group",
                title="Perbandingan Budget vs Serapan Dana per SSR",
                xaxis_title="SSR",
                yaxis_title="Nominal Dana (Rp)",
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                height=380,
                margin=dict(l=20, r=20, t=40, b=20)
            )
            st.plotly_chart(fig_dana, use_container_width=True)

        with col_t_dana:
            st.markdown("**Tabel Serapan Dana per SSR:**")
            df_t_dana = df_ssr[["ssr", "tot_real_dana", "tot_plan_dana", "% Serapan Dana"]].copy()
            df_t_dana.columns = ["SSR", "Realisasi (Rp)", "Budget (Rp)", "% Serapan"]
            df_t_dana = df_t_dana.sort_values("% Serapan", ascending=False)
            st.dataframe(
                df_t_dana.style.format({
                    "Realisasi (Rp)": "Rp {:,.0f}",
                    "Budget (Rp)": "Rp {:,.0f}",
                    "% Serapan": "{:.2f}%"
                }),
                use_container_width=True,
                hide_index=True,
                height=340
            )

    with tab_ssr_keg:
        col_c_keg, col_t_keg = st.columns([6, 4])
        with col_c_keg:
            fig_keg = go.Figure()
            fig_keg.add_trace(go.Bar(
                x=df_ssr["ssr"],
                y=df_ssr["tot_plan_keg"],
                name="Target Kegiatan",
                marker_color="#94a3b8"
            ))
            fig_keg.add_trace(go.Bar(
                x=df_ssr["ssr"],
                y=df_ssr["tot_real_keg"],
                name="Kegiatan Terlaksana",
                marker_color="#3b82f6"
            ))
            fig_keg.update_layout(
                barmode="group",
                title="Perbandingan Kegiatan Terlaksana vs Target per SSR",
                xaxis_title="SSR",
                yaxis_title="Jumlah Kegiatan",
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                height=380,
                margin=dict(l=20, r=20, t=40, b=20)
            )
            st.plotly_chart(fig_keg, use_container_width=True)

        with col_t_keg:
            st.markdown("**Tabel Capaian Kegiatan per SSR:**")
            df_t_keg = df_ssr[["ssr", "tot_real_keg", "tot_plan_keg", "% Serapan Kegiatan"]].copy()
            df_t_keg.columns = ["SSR", "Terlaksana (Keg)", "Target (Keg)", "% Capaian"]
            df_t_keg = df_t_keg.sort_values("% Capaian", ascending=False)
            st.dataframe(
                df_t_keg.style.format({
                    "Terlaksana (Keg)": "{:,.0f}",
                    "Target (Keg)": "{:,.0f}",
                    "% Capaian": "{:.2f}%"
                }),
                use_container_width=True,
                hide_index=True,
                height=340
            )

    st.markdown("<br>", unsafe_allow_html=True)

    # ==========================================================================
    # 4. RANK 10 BL DENGAN SERAPAN PALING BAGUS (TERTINGGI)
    # ==========================================================================
    st.subheader("🏆 Peringkat 10 BL dengan Serapan Paling Bagus")

    df_bl = df_filtered.groupby("bl").agg(
        tot_plan_dana=("total_budget_valid", "sum"),
        tot_real_dana=("realisasi", "sum"),
        tot_plan_keg=("jml_kegiatan_planning", "sum"),
        tot_real_keg=("jml_kegiatan_realisasi", "sum")
    ).reset_index()

    df_bl_active = df_bl[df_bl["tot_plan_dana"] > 0].copy()
    df_bl_active["% Serapan"] = (df_bl_active["tot_real_dana"] / df_bl_active["tot_plan_dana"]) * 100
    df_bl_active["BL_Label"] = "BL " + df_bl_active["bl"].astype(int).astype(str)

    top_10_bl = df_bl_active.sort_values("% Serapan", ascending=False).head(10)
    bottom_10_bl = df_bl_active.sort_values("% Serapan", ascending=True).head(10)

    col_top10, col_bot10 = st.columns(2)

    with col_top10:
        st.markdown("##### 🥇 Top 10 BL dengan Serapan Paling Bagus (Tertinggi)")
        fig_top = px.bar(
            top_10_bl,
            x="% Serapan",
            y="BL_Label",
            orientation="h",
            text=top_10_bl["% Serapan"].apply(lambda x: f"{x:.1f}%"),
            color="% Serapan",
            color_continuous_scale="Greens"
        )
        fig_top.update_layout(
            yaxis=dict(autorange="reversed"),
            xaxis_title="% Serapan Anggaran",
            yaxis_title="Budget Line",
            height=380,
            margin=dict(l=20, r=20, t=20, b=20)
        )
        st.plotly_chart(fig_top, use_container_width=True)

        with st.expander("Lihat Rincian Top 10 BL"):
            st.dataframe(
                top_10_bl[["BL_Label", "tot_real_dana", "tot_plan_dana", "% Serapan"]].rename(
                    columns={"BL_Label": "Budget Line", "tot_real_dana": "Realisasi (Rp)", "tot_plan_dana": "Budget (Rp)"}
                ).style.format({
                    "Realisasi (Rp)": "Rp {:,.0f}",
                    "Budget (Rp)": "Rp {:,.0f}",
                    "% Serapan": "{:.2f}%"
                }),
                use_container_width=True,
                hide_index=True
            )

    with col_bot10:
        st.markdown("##### 🔻 Bottom 10 BL (Serapan Terendah untuk Evaluasi)")
        fig_bot = px.bar(
            bottom_10_bl,
            x="% Serapan",
            y="BL_Label",
            orientation="h",
            text=bottom_10_bl["% Serapan"].apply(lambda x: f"{x:.1f}%"),
            color="% Serapan",
            color_continuous_scale="Reds_r"
        )
        fig_bot.update_layout(
            yaxis=dict(autorange="reversed"),
            xaxis_title="% Serapan Anggaran",
            yaxis_title="Budget Line",
            height=380,
            margin=dict(l=20, r=20, t=20, b=20)
        )
        st.plotly_chart(fig_bot, use_container_width=True)

        with st.expander("Lihat Rincian Bottom 10 BL"):
            st.dataframe(
                bottom_10_bl[["BL_Label", "tot_real_dana", "tot_plan_dana", "% Serapan"]].rename(
                    columns={"BL_Label": "Budget Line", "tot_real_dana": "Realisasi (Rp)", "tot_plan_dana": "Budget (Rp)"}
                ).style.format({
                    "Realisasi (Rp)": "Rp {:,.0f}",
                    "Budget (Rp)": "Rp {:,.0f}",
                    "% Serapan": "{:.2f}%"
                }),
                use_container_width=True,
                hide_index=True
            )

    st.markdown("<br>", unsafe_allow_html=True)

    # ==========================================================================
    # 5. TABEL DETAIL & FITUR UNDUH DATA
    # ==========================================================================
    st.markdown("---")
    st.subheader("📑 Detail Data Aktivitas & Justifikasi")

    cols_show = [
        "ssr", "bulan_nama", "bl", "jml_kegiatan_planning", "jml_kegiatan_realisasi",
        "total_budget_valid", "realisasi", "justifikasi"
    ]
    avail_cols = [c for c in cols_show if c in df_filtered.columns]

    df_display = df_filtered[avail_cols].copy()
    df_display.columns = [
        "SSR", "Bulan", "BL", "Planning (Keg)", "Realisasi (Keg)",
        "Total Budget (Rp)", "Realisasi (Rp)", "Justifikasi"
    ]

    st.dataframe(
        df_display.sort_values(["SSR", "BL"]),
        use_container_width=True,
        hide_index=True,
        height=320
    )

    # Tombol Download CSV
    csv_bytes = df_filtered.to_csv(index=False).encode("utf-8")
    st.download_button(
        label="📥 Unduh Data Terfilter (CSV)",
        data=csv_bytes,
        file_name="Monitoring_Kegiatan_Anggaran_Filtered.csv",
        mime="text/csv"
    )


# --- RENDER DASHBOARD UTAMA ---
st.title("📊 Dashboard Monitoring Kegiatan & Realisasi Anggaran")
render_interactive_dashboard(df_clean)
