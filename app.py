import warnings
warnings.filterwarnings("ignore")

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path
import io
from sklearn.preprocessing import StandardScaler
from sklearn.cluster import KMeans
from sklearn.metrics import silhouette_score

# --- KONFIGURASI HALAMAN ---
st.set_page_config(
    page_title="Dashboard Monitoring Kegiatan & Anggaran",
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


@st.cache_data
def get_bl_metadata():
    """Membaca file Data BL Rev.xlsx untuk memetakan Group Budget dan Deskripsi Aktivitas"""
    for f_path in ["Data BL Rev.xlsx", "data/Data BL Rev.xlsx"]:
        if Path(f_path).exists():
            try:
                df_meta = pd.read_excel(f_path, sheet_name="Group Budget")
                df_meta = df_meta.rename(columns={
                    "BL": "bl",
                    "Description of Activities": "bl_desc",
                    "Group Budget": "group_budget"
                })
                df_meta["bl"] = pd.to_numeric(df_meta["bl"], errors="coerce")
                df_meta["group_budget"] = df_meta["group_budget"].fillna("Lainnya").astype(str).str.strip()
                df_meta["bl_desc"] = df_meta["bl_desc"].fillna("-").astype(str).str.strip()
                return df_meta[["bl", "group_budget", "bl_desc"]].dropna(subset=["bl"]).drop_duplicates(subset=["bl"])
            except Exception:
                pass
    return None


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

    # Group Budget: prioritaskan kolom yang SUDAH ADA di Excel terbaru (mis. "Grup budget").
    # Ini penting agar data ALL BL tetap utuh, sementara 17 BL kategori Kegiatan tetap bisa difilter dengan benar.
    group_col = None
    for c in df.columns:
        c_low = str(c).strip().lower()
        if c_low in ["grup budget", "group budget", "group_budget", "kategori budget", "kelompok budget"]:
            group_col = c
            break

    if group_col is not None:
        df["group_budget"] = df[group_col].fillna("Lainnya").astype(str).str.strip()
        # Normalisasi ejaan kategori agar pie chart tidak kehilangan kategori karena spasi/case.
        gb_norm = {
            "sdm": "SDM",
            "salary": "SDM",
            "kegiatan": "Kegiatan",
            "activity": "Kegiatan",
            "aset": "Aset",
            "asset": "Aset",
            "operasional": "Operasional",
            "management/fixed cost": "Operasional",
            "management fixed cost": "Operasional",
            "performance based": "Performance Based",
            "pb pl": "Performance Based",
        }
        df["group_budget"] = df["group_budget"].apply(lambda x: gb_norm.get(str(x).strip().lower(), str(x).strip()))
        if group_col != "group_budget":
            df = df.drop(columns=[group_col])
        # Deskripsi aktivitas tetap dicoba dari metadata eksternal jika tersedia, tanpa mengubah group budget dari Excel utama.
        df_bl_meta = get_bl_metadata()
        if df_bl_meta is not None:
            meta_desc = df_bl_meta[["bl", "bl_desc"]].drop_duplicates(subset=["bl"])
            df = df.merge(meta_desc, on="bl", how="left")
            df["bl_desc"] = df["bl_desc"].fillna("-")
        else:
            df["bl_desc"] = "-"
    else:
        # Fallback untuk file lama yang belum punya kolom Group Budget.
        df_bl_meta = get_bl_metadata()
        if df_bl_meta is not None:
            df = df.merge(df_bl_meta, on="bl", how="left")
            df["group_budget"] = df["group_budget"].fillna("Lainnya")
            df["bl_desc"] = df["bl_desc"].fillna("-")
        else:
            # Fallback mapping 5 kategori berdasarkan Dashboard baru.xlsx terbaru.
            # Dipakai bila file yang dijalankan belum memiliki kolom "Grup budget" / metadata eksternal.
            group_map = {
                # SDM
                **{bl: "SDM" for bl in [5, 6, 7, 8, 9, 10, 11, 12, 41, 42, 43, 44, 45, 46, 47, 48, 50, 51, 52, 53, 90, 91, 92, 93, 146, 148]},
                # Kegiatan
                **{bl: "Kegiatan" for bl in [36, 37, 38, 39, 40, 49, 54, 55, 56, 61, 62, 68, 69, 74, 78, 80, 154]},
                # Aset
                153: "Aset",
                # Operasional
                **{bl: "Operasional" for bl in [19, 20, 34, 60, 70, 140]},
                # Performance Based
                **{bl: "Performance Based" for bl in [57, 58, 63, 66, 88, 89, 155]},
            }
            df["group_budget"] = df["bl"].map(group_map).fillna("Lainnya")
            df["bl_desc"] = "-"

    return df.reset_index(drop=True)


# ==============================================================================
# SIDEBAR: SUMBER DATA & SMART HEADER TRACER
# ==============================================================================
st.sidebar.markdown("###  Sumber Data Excel")

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
        st.error("File Excel tidak ditemukan. Silakan unggah file Excel pada menu di sidebar.")
        st.stop()

# 2. Pemilihan Sheet Dinamis
available_sheets = read_excel_file(file_bytes)
default_sheet_index = available_sheets.index("Activity") if "Activity" in available_sheets else 0

selected_sheet = st.sidebar.selectbox(
    "Pilih Sheet Data:",
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
with st.sidebar.expander("Modifikasi & Tracing Kolom", expanded=False):
    st.markdown(f"""
    <div class="trace-card">
        <b>Status Auto-Trace Sheet:</b><br>
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
    st.sidebar.markdown("### Filter Analisis")
    
    # 1. Filter Fokus Group Budget (Default: Khusus Kegiatan (17 BL))
    group_options = ["Khusus Kegiatan Saja (17 BL)", "Semua Group Budget"]
    if "group_budget" in df_base.columns:
        other_groups = [g for g in sorted(df_base["group_budget"].unique()) if g not in ["Lainnya", "-", "Kegiatan"]]
        group_options += other_groups

    selected_group_focus = st.sidebar.selectbox(
        "Fokus Kategori (Group Budget):",
        options=group_options,
        index=0,
        help="Sesuai arahan rapat, visualisasi default berfokus ke 17 BL kategori Kegiatan."
    )

    # 2. Filter Periode Bulan
    available_months = df_base.sort_values("bulan_no")["bulan_nama"].dropna().unique().tolist()
    selected_months = st.sidebar.multiselect(
        "Pilih Periode Bulan:",
        options=available_months,
        default=available_months,
        help="Pilih satu atau lebih bulan"
    )

    # 3. Filter SSR
    available_ssr = sorted(df_base["ssr"].dropna().unique().tolist())
    selected_ssr = st.sidebar.multiselect(
        "Pilih SSR (Sub-Sub Recipient):",
        options=available_ssr,
        default=available_ssr
    )

    # Filter dasar berdasarkan Group Budget
    if selected_group_focus == "Khusus Kegiatan Saja (17 BL)":
        df_group_filtered = df_base[df_base["group_budget"] == "Kegiatan"].copy()
    elif selected_group_focus == "Semua Group Budget":
        df_group_filtered = df_base.copy()
    else:
        df_group_filtered = df_base[df_base["group_budget"] == selected_group_focus].copy()

    # 4. Filter Budget Line (BL) dinamis sesuai Group Budget terpilih
    available_bl = sorted(df_group_filtered["bl"].dropna().unique().astype(int).tolist())
    select_all_bl_toggle = st.sidebar.checkbox(f"Pilih Semua Budget Line ({len(available_bl)} BL)", value=True)
    
    if select_all_bl_toggle:
        selected_bl = available_bl
    else:
        selected_bl = st.sidebar.multiselect(
            "Pilih BL Tertentu:",
            options=available_bl,
            default=available_bl[:10] if len(available_bl) >= 10 else available_bl
        )

    # Dataset SEMUA BL untuk KPI utama, tren bulanan, dan agregasi SSR
    # Tetap mengikuti filter Bulan & SSR, tetapi TIDAK dibatasi Group Budget/BL kegiatan.
    df_all_filtered = df_base[
        (df_base["bulan_nama"].isin(selected_months)) &
        (df_base["ssr"].isin(selected_ssr))
    ].copy()

    # Dataset terfilter sesuai fokus Group Budget/BL untuk analisis khusus kegiatan
    df_filtered = df_group_filtered[
        (df_group_filtered["bulan_nama"].isin(selected_months)) &
        (df_group_filtered["ssr"].isin(selected_ssr)) &
        (df_group_filtered["bl"].isin(selected_bl))
    ].copy()

    # Info Sumber Data & Status Fokus
    badge_group = "Khusus Kegiatan (17 BL)" if selected_group_focus == "Khusus Kegiatan Saja (17 BL)" else selected_group_focus
    st.markdown(f"**Sumber Data:** `{file_source_name}` (Sheet: `{selected_sheet}`) | **Kategori:** `{badge_group}` | **Data Ditampilkan:** `{len(df_filtered):,}` baris")

    if df_filtered.empty:
        st.warning("Tidak ada data yang sesuai dengan kombinasi filter yang dipilih.")
        return

    st.markdown("---")

    # ==========================================================================
    # 1. MAIN KPI: JUMLAH KEGIATAN TERLAKSANA & BUDGET BESERTA PERSENTASENYA
    # ==========================================================================
    st.subheader("Main KPI (Indikator Utama)")
    tot_keg_plan = df_all_filtered["jml_kegiatan_planning"].sum()
    tot_keg_real = df_all_filtered["jml_kegiatan_realisasi"].sum()
    persen_kegiatan = (tot_keg_real / tot_keg_plan * 100) if tot_keg_plan > 0 else 0

    tot_dana_plan = df_all_filtered["total_budget_valid"].sum()
    tot_dana_real = df_all_filtered["realisasi"].sum()
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

    # ==========================================================================
    # KOMPOSISI 5 KATEGORI GROUP BUDGET (ALL BL) - PIE CHART BUDGET ONLY
    # ==========================================================================
    # Pie chart hanya menunjukkan KOMPOSISI TOTAL BUDGET.
    # Tidak menggunakan realisasi maupun % serapan dalam perhitungan/tooltip.
    # Sumber kategori berasal dari seluruh BL (ALL BL), bukan hanya 17 BL Kegiatan.
    if "group_budget" in df_all_filtered.columns:
        kategori_utama = ["SDM", "Kegiatan", "Aset", "Operasional", "Performance Based"]

        df_gb_budget = (
            df_all_filtered[df_all_filtered["group_budget"].isin(kategori_utama)]
            .groupby("group_budget", as_index=False)
            .agg(budget=("total_budget_valid", "sum"))
        )

        # Pastikan 5 kategori tetap tersedia walaupun salah satu bernilai 0.
        df_gb_budget = (
            pd.DataFrame({"group_budget": kategori_utama})
            .merge(df_gb_budget, on="group_budget", how="left")
            .fillna({"budget": 0})
        )

        st.subheader("Komposisi Budget berdasarkan 5 Kategori BL")
        st.caption("Pie chart menunjukkan proporsi total budget ALL BL untuk kategori SDM, Kegiatan, Aset, Operasional, dan Performance Based.")

        fig_gb_pie = px.pie(
            df_gb_budget,
            names="group_budget",
            values="budget",
            hole=0.38,
            title="Proporsi Total Budget per Kategori"
        )
        # Buat semua kategori tetap terbaca, termasuk kategori dengan proporsi sangat kecil seperti Aset
        total_budget_pie = df_gb_budget["budget"].sum()
        df_gb_budget["persen"] = np.where(
            total_budget_pie > 0,
            df_gb_budget["budget"] / total_budget_pie * 100,
            0
        )

        # Susun custom label: kategori kecil tetap punya label di luar donut
        custom_text = [
            f"{row['group_budget']}<br>{row['persen']:.2f}%"
            for _, row in df_gb_budget.iterrows()
        ]

        fig_gb_pie = px.pie(
            df_gb_budget,
            names="group_budget",
            values="budget",
            hole=0.48,
            title="Proporsi Total Budget per Kategori"
        )
        fig_gb_pie.update_traces(
            text=custom_text,
            textinfo="text",
            textposition="outside",
            pull=[0.00, 0.00, 0.08, 0.00, 0.00],
            automargin=True,
            hovertemplate=(
                "<b>%{label}</b><br>"
                "Budget: Rp %{value:,.0f}<br>"
                "Proporsi: %{percent}<extra></extra>"
            )
        )
        fig_gb_pie.update_layout(
            height=460,
            margin=dict(l=90, r=160, t=65, b=55),
            uniformtext_minsize=10,
            uniformtext_mode="show",
            legend=dict(
                orientation="v",
                yanchor="middle",
                y=0.5,
                xanchor="left",
                x=1.02
            )
        )
        st.plotly_chart(fig_gb_pie, use_container_width=True)

        # Ringkasan nominal membantu memastikan seluruh 5 kategori terlihat meski slice sangat kecil
        df_gb_show = df_gb_budget.copy()
        df_gb_show["Proporsi"] = df_gb_show["persen"]
        df_gb_show = df_gb_show[["group_budget", "budget", "Proporsi"]]
        df_gb_show.columns = ["Kategori", "Total Budget (Rp)", "Proporsi"]
        st.dataframe(
            df_gb_show.style.format({
                "Total Budget (Rp)": "Rp {:,.0f}",
                "Proporsi": "{:.2f}%"
            }),
            use_container_width=True,
            hide_index=True
        )

    st.markdown("<br>", unsafe_allow_html=True)

    # ==========================================================================
    # 2. DETAIL BUDGET VS SERAPAN PER BULAN
    # ==========================================================================
    st.subheader("Detail Budget vs Serapan per Bulan")
    
    df_monthly = df_all_filtered.groupby(["bulan_no", "bulan_nama"]).agg(
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
    # 3. BUDGET VS SERAPAN DANA PER SSR (ALL BL)
    # ==========================================================================
    # Hanya menampilkan Budget vs Serapan Dana per SSR.
    # Chart "Kegiatan Terlaksana vs Target per SSR" sengaja dihapus sesuai revisi.
    st.subheader("Budget vs Serapan Dana per SSR")

    df_ssr_display = df_all_filtered.groupby("ssr").agg(
        tot_plan_dana=("total_budget_valid", "sum"),
        tot_real_dana=("realisasi", "sum")
    ).reset_index()

    df_ssr_display["% Serapan Dana"] = np.where(
        df_ssr_display["tot_plan_dana"] > 0,
        df_ssr_display["tot_real_dana"] / df_ssr_display["tot_plan_dana"] * 100,
        0
    )

    col_c_dana, col_t_dana = st.columns([6, 4])
    with col_c_dana:
        fig_dana = go.Figure()
        fig_dana.add_trace(go.Bar(
            x=df_ssr_display["ssr"],
            y=df_ssr_display["tot_plan_dana"],
            name="Total Budget",
            marker_color="#cbd5e1"
        ))
        fig_dana.add_trace(go.Bar(
            x=df_ssr_display["ssr"],
            y=df_ssr_display["tot_real_dana"],
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
        df_t_dana = df_ssr_display[["ssr", "tot_real_dana", "tot_plan_dana", "% Serapan Dana"]].copy()
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

    st.markdown("<br>", unsafe_allow_html=True)

    # ==========================================================================
    # 4. SEGMENTASI KINERJA SSR DENGAN K-MEANS CLUSTERING (MACHINE LEARNING)
    # ==========================================================================
    st.subheader("Segmentasi Kinerja Lembaga (K-Means Clustering)")
    # K-Means khusus 17 BL Kegiatan, tetap mengikuti filter Bulan & SSR.
    kegiatan_17_bl = [36, 37, 38, 39, 40, 49, 54, 55, 56, 61, 62, 68, 69, 74, 78, 80, 154]
    df_kmeans = df_base[
        (df_base["bulan_nama"].isin(selected_months)) &
        (df_base["ssr"].isin(selected_ssr)) &
        (df_base["bl"].isin(kegiatan_17_bl))
    ].copy()

    df_ssr = df_kmeans.groupby("ssr").agg(
        tot_plan_keg=("jml_kegiatan_planning", "sum"),
        tot_real_keg=("jml_kegiatan_realisasi", "sum"),
        tot_plan_dana=("total_budget_valid", "sum"),
        tot_real_dana=("realisasi", "sum")
    ).reset_index()
    df_ssr["% Serapan Dana"] = np.where(df_ssr["tot_plan_dana"] > 0, df_ssr["tot_real_dana"] / df_ssr["tot_plan_dana"] * 100, 0)
    df_ssr["% Serapan Kegiatan"] = np.where(df_ssr["tot_plan_keg"] > 0, df_ssr["tot_real_keg"] / df_ssr["tot_plan_keg"] * 100, 0)

    st.markdown("""
    Penerapan algoritma **K-Means Clustering (*Unsupervised Machine Learning*)** untuk mengelompokkan 
    **Sub-Sub Recipient (SSR)** ke dalam klaster performa objektif berdasarkan indikator **% Capaian Kegiatan** dan **% Serapan Anggaran**.
    """)

    if len(df_ssr) >= 3:
        X_cluster = df_ssr[["% Serapan Kegiatan", "% Serapan Dana"]].values
        scaler = StandardScaler()
        X_scaled = scaler.fit_transform(X_cluster)

        max_k = min(6, len(df_ssr) - 1)
        k_list = list(range(2, max_k + 1))
        inertias = []
        silhouettes = []
        for k_val in k_list:
            km_test = KMeans(n_clusters=k_val, random_state=42, n_init=10)
            km_test.fit(X_scaled)
            inertias.append(km_test.inertia_)
            silhouettes.append(silhouette_score(X_scaled, km_test.labels_))

        col_ctrl1, col_ctrl2, col_ctrl3 = st.columns([3, 3, 4])
        with col_ctrl1:
            chosen_k = st.selectbox("Pilih Jumlah Klaster (K):", options=k_list, index=k_list.index(3) if 3 in k_list else 0)
        
        km_final = KMeans(n_clusters=chosen_k, random_state=42, n_init=10)
        df_ssr["Cluster_ID"] = km_final.fit_predict(X_scaled)
        current_sil = silhouette_score(X_scaled, df_ssr["Cluster_ID"])

        with col_ctrl2:
            st.metric("Silhouette Score (Evaluasi)", f"{current_sil:.3f}", "Sangat Baik (>0.5)" if current_sil > 0.5 else "Cukup")
        with col_ctrl3:
            st.metric("Metode Evaluasi K Optimal", "Elbow & Silhouette", "K=3 Paling Optimal")

        col_eval, col_scatter = st.columns([5, 5])

        with col_eval:
            fig_eval = go.Figure()
            fig_eval.add_trace(go.Scatter(
                x=k_list, y=inertias, mode="lines+markers", name="Inertia (Elbow)",
                line=dict(color="#f59e0b", width=2), yaxis="y1"
            ))
            fig_eval.add_trace(go.Scatter(
                x=k_list, y=silhouettes, mode="lines+markers", name="Silhouette Score",
                line=dict(color="#10b981", width=2, dash="dot"), yaxis="y2"
            ))
            fig_eval.update_layout(
                title="Evaluasi K Optimal (Elbow vs Silhouette)",
                xaxis=dict(title=dict(text="Jumlah Klaster (K)"), tickmode="linear"),
                yaxis=dict(title=dict(text="Inertia (WCSS)", font=dict(color="#f59e0b"))),
                yaxis2=dict(title=dict(text="Silhouette Score", font=dict(color="#10b981")), overlaying="y", side="right"),
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                height=380,
                margin=dict(l=20, r=20, t=40, b=20)
            )
            st.plotly_chart(fig_eval, use_container_width=True)

        with col_scatter:
            df_ssr["Label_Klaster"] = "Klaster " + df_ssr["Cluster_ID"].astype(str)
            fig_cluster = px.scatter(
                df_ssr,
                x="% Serapan Kegiatan",
                y="% Serapan Dana",
                color="Label_Klaster",
                text="ssr",
                hover_name="ssr",
                hover_data={"tot_real_keg": ":,.0f", "tot_real_dana": ":,.0f", "% Serapan Kegiatan": ":.1f", "% Serapan Dana": ":.1f", "Label_Klaster": False},
                title="Visualisasi Segmentasi Kinerja SSR (2D Scatter Plot)",
                color_discrete_sequence=px.colors.qualitative.Set2
            )
            fig_cluster.update_traces(textposition="top center", marker=dict(size=14, line=dict(width=1, color="DarkSlateGrey")))
            fig_cluster.update_layout(
                xaxis_title="% Capaian Kegiatan",
                yaxis_title="% Serapan Anggaran",
                height=380,
                margin=dict(l=20, r=20, t=40, b=20)
            )
            st.plotly_chart(fig_cluster, use_container_width=True)

        st.markdown("##### Profiling Klaster & Rekomendasi Manajerial")
        cluster_summary = []
        for cid in sorted(df_ssr["Cluster_ID"].unique()):
            sub_c = df_ssr[df_ssr["Cluster_ID"] == cid]
            avg_keg = sub_c["% Serapan Kegiatan"].mean()
            avg_dana = sub_c["% Serapan Dana"].mean()
            anggota = ", ".join(sub_c["ssr"].tolist())
            
            if avg_keg >= 75 and 75 <= avg_dana <= 110:
                status_profil = "Kinerja Prima (Stabil)"
                rekomendasi = "Pertahankan performa dan jadikan acuan bagi SSR lainnya."
            elif avg_keg < 50 and avg_dana < 50:
                status_profil = "Serapan & Kegiatan Rendah (Kritis)"
                rekomendasi = "Perlu evaluasi kendala teknis lapangan dan intervensi khusus."
            elif avg_keg > 150:
                status_profil = "Over-Achieving Kegiatan"
                rekomendasi = "Apresiasi kinerja & validasi kecukupan alokasi anggaran kegiatan."
            elif avg_dana > 110:
                status_profil = "Risiko Overbudget"
                rekomendasi = "Lakukan audit pengendalian anggaran agar tidak defisit."
            else:
                status_profil = "Kinerja Menengah"
                rekomendasi = "Tingkatkan akselerasi kegiatan di sisa periode berjalan."

            cluster_summary.append({
                "Klaster": f"Klaster {cid}",
                "Status Profiling": status_profil,
                "Jumlah SSR": len(sub_c),
                "Rata-rata Capaian Kegiatan": f"{avg_keg:.1f}%",
                "Rata-rata Serapan Anggaran": f"{avg_dana:.1f}%",
                "Anggota SSR": anggota,
                "Rekomendasi Manajerial": rekomendasi
            })

        df_summary_cluster = pd.DataFrame(cluster_summary)
        st.dataframe(df_summary_cluster, use_container_width=True, hide_index=True)
    else:
        st.info("Data SSR tidak mencukupi untuk analisis K-Means Clustering.")

    st.markdown("<br>", unsafe_allow_html=True)

    # ==========================================================================
    # 5. PANEL TOP 10 BERDAMPINGAN (KEGIATAN/BL DI KIRI & SSR DI KANAN)
    # ==========================================================================
    st.subheader("Panel Top 10 Berdampingan (Sesuai Notulensi Rapat)")
    st.markdown("Menyandingkan **Top 10 Kegiatan/BL** (sisi kiri) dan **Top 10 SSR Agregat** (sisi kanan) dengan sumbu **Persentase Serapan (%)**.")

    # 1. Dataset khusus 17 BL Kegiatan (sesuai notulensi/versi sebelumnya)
    kegiatan_17_bl = [36, 37, 38, 39, 40, 49, 54, 55, 56, 61, 62, 68, 69, 74, 78, 80, 154]
    df_top10 = df_base[
        (df_base["bulan_nama"].isin(selected_months)) &
        (df_base["ssr"].isin(selected_ssr)) &
        (df_base["bl"].isin(kegiatan_17_bl))
    ].copy()

    # 1. Agregasi BL
    agg_dict = {
        "tot_plan_dana": ("total_budget_valid", "sum"),
        "tot_real_dana": ("realisasi", "sum"),
        "tot_plan_keg": ("jml_kegiatan_planning", "sum"),
        "tot_real_keg": ("jml_kegiatan_realisasi", "sum")
    }
    if "bl_desc" in df_top10.columns:
        agg_dict["bl_desc"] = ("bl_desc", "first")
    if "group_budget" in df_top10.columns:
        agg_dict["group_budget"] = ("group_budget", "first")

    df_bl = df_top10.groupby("bl").agg(**agg_dict).reset_index()
    df_bl_active = df_bl[df_bl["tot_plan_dana"] > 0].copy()
    df_bl_active["% Serapan"] = (df_bl_active["tot_real_dana"] / df_bl_active["tot_plan_dana"]) * 100
    df_bl_active["BL_Label"] = "BL " + df_bl_active["bl"].astype(int).astype(str)
    if "bl_desc" not in df_bl_active.columns:
        df_bl_active["bl_desc"] = "-"

    top_10_bl = df_bl_active.sort_values("% Serapan", ascending=False).head(10)

    # 2. Agregasi SSR
    df_ssr_top = df_top10.groupby("ssr").agg(
        tot_plan_dana=("total_budget_valid", "sum"),
        tot_real_dana=("realisasi", "sum")
    ).reset_index()
    df_ssr_top["% Serapan Total"] = np.where(df_ssr_top["tot_plan_dana"] > 0, df_ssr_top["tot_real_dana"] / df_ssr_top["tot_plan_dana"] * 100, 0)
    top_10_ssr = df_ssr_top.sort_values("% Serapan Total", ascending=False).head(10)

    col_top_bl, col_top_ssr = st.columns(2)

    with col_top_bl:
        st.markdown("##### Panel Kiri: Top 10 Kegiatan / BL")
        fig_top_bl = px.bar(
            top_10_bl,
            x="% Serapan",
            y="BL_Label",
            orientation="h",
            text=top_10_bl["% Serapan"].apply(lambda x: f"{x:.1f}%"),
            hover_data={"bl_desc": True, "% Serapan": ":.2f", "tot_real_dana": ":,.0f", "tot_plan_dana": ":,.0f", "BL_Label": False},
            labels={"tot_real_dana": "Realisasi (Rp)", "tot_plan_dana": "Budget (Rp)", "bl_desc": "Nama Kegiatan"},
            color="% Serapan",
            color_continuous_scale="Greens"
        )
        fig_top_bl.update_layout(
            yaxis=dict(autorange="reversed"),
            xaxis_title="% Serapan Anggaran",
            yaxis_title="Budget Line (BL)",
            height=380,
            margin=dict(l=20, r=20, t=20, b=20)
        )
        st.plotly_chart(fig_top_bl, use_container_width=True)

    with col_top_ssr:
        st.markdown("##### Panel Kanan: Top 10 SSR (Agregat Total)")
        fig_top_ssr = px.bar(
            top_10_ssr,
            x="% Serapan Total",
            y="ssr",
            orientation="h",
            text=top_10_ssr["% Serapan Total"].apply(lambda x: f"{x:.1f}%"),
            hover_data={"tot_real_dana": ":,.0f", "tot_plan_dana": ":,.0f", "% Serapan Total": ":.2f", "ssr": False},
            labels={"tot_real_dana": "Realisasi (Rp)", "tot_plan_dana": "Budget (Rp)"},
            color="% Serapan Total",
            color_continuous_scale="Blues"
        )
        fig_top_ssr.update_layout(
            yaxis=dict(autorange="reversed"),
            xaxis_title="% Serapan Total",
            yaxis_title="Sub-Sub Recipient (SSR)",
            height=380,
            margin=dict(l=20, r=20, t=20, b=20)
        )
        st.plotly_chart(fig_top_ssr, use_container_width=True)

    # Expander: Visualisasi Lengkap 17 BL Kegiatan & Evaluasi Serapan Terendah
    with st.expander("Lihat Visualisasi Lengkap 17 BL Kategori Kegiatan & Evaluasi"):
        tab_keg_all, tab_bot_bl = st.tabs(["Seluruh BL Kategori Kegiatan", "Evaluasi Serapan Terendah"])
        
        with tab_keg_all:
            if "group_budget" in df_bl_active.columns:
                df_keg_17 = df_bl_active[df_bl_active["group_budget"] == "Kegiatan"].sort_values("% Serapan", ascending=False)
            else:
                df_keg_17 = df_bl_active.sort_values("% Serapan", ascending=False)
                
            fig_keg_17 = px.bar(
                df_keg_17,
                x="% Serapan",
                y="BL_Label",
                orientation="h",
                text=df_keg_17["% Serapan"].apply(lambda x: f"{x:.1f}%"),
                hover_data={"bl_desc": True, "tot_real_dana": ":,.0f", "tot_plan_dana": ":,.0f", "% Serapan": ":.2f", "BL_Label": False},
                labels={"tot_real_dana": "Realisasi (Rp)", "tot_plan_dana": "Budget (Rp)", "bl_desc": "Aktivitas"},
                color="% Serapan",
                color_continuous_scale="Teal"
            )
            fig_keg_17.update_layout(
                yaxis=dict(autorange="reversed"),
                xaxis_title="% Serapan Anggaran",
                yaxis_title="Budget Line Kegiatan",
                height=450,
                margin=dict(l=20, r=20, t=20, b=20)
            )
            st.plotly_chart(fig_keg_17, use_container_width=True)
            
            st.dataframe(
                df_keg_17[["BL_Label", "bl_desc", "tot_real_dana", "tot_plan_dana", "% Serapan"]].rename(
                    columns={"BL_Label": "Kode BL", "bl_desc": "Deskripsi Kegiatan", "tot_real_dana": "Realisasi (Rp)", "tot_plan_dana": "Budget (Rp)"}
                ).style.format({
                    "Realisasi (Rp)": "Rp {:,.0f}",
                    "Budget (Rp)": "Rp {:,.0f}",
                    "% Serapan": "{:.2f}%"
                }),
                use_container_width=True,
                hide_index=True
            )

        with tab_bot_bl:
            bottom_10_bl = df_bl_active.sort_values("% Serapan", ascending=True).head(10)
            fig_bot = px.bar(
                bottom_10_bl,
                x="% Serapan",
                y="BL_Label",
                orientation="h",
                text=bottom_10_bl["% Serapan"].apply(lambda x: f"{x:.1f}%"),
                hover_data={"bl_desc": True, "tot_real_dana": ":,.0f", "tot_plan_dana": ":,.0f", "% Serapan": ":.2f", "BL_Label": False},
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

    st.markdown("<br>", unsafe_allow_html=True)

    # ==========================================================================
    # 5. TABEL DETAIL & FITUR UNDUH DATA
    # ==========================================================================
    st.markdown("---")
    st.subheader("Detail Data Aktivitas & Justifikasi")

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
        label="Unduh Data Terfilter (CSV)",
        data=csv_bytes,
        file_name="Monitoring_Kegiatan_Anggaran_Filtered.csv",
        mime="text/csv"
    )


# --- RENDER DASHBOARD UTAMA ---
st.title("Dashboard Monitoring Kegiatan & Realisasi Anggaran")
render_interactive_dashboard(df_clean)
