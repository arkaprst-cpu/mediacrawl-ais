# ══════════════════════════════════════════════════════════════════════════
# HALAMAN 3 — REPOSITORI ISU (Langkah Kerja 4)
# ══════════════════════════════════════════════════════════════════════════
# Menampilkan hasil telaah matang (Status Review = "Sudah Direview") yang
# sudah di-export sebagai Excel dan ditaruh MANUAL oleh tim ke folder
# Google Drive tertentu. Halaman ini hanya MEMBACA folder itu — tidak ada
# tulis-menulis ke Drive dari app. Navigasi utama: Sektor → Tema → Topik.

import streamlit as st
import pandas as pd
import io
import html as html_lib
from struktur_app import STRUKTUR_APP

st.markdown("""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;600&display=swap');
  html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

  /* .main-header didefinisikan ULANG di sini (bukan cuma mengandalkan CSS
     dari blok Crawl Berita di app.py) karena halaman ini punya 2 jalur
     akses: (1) lewat nav sidebar setelah login — di jalur ini CSS dari
     Crawl Berita kebetulan sudah kepasang duluan karena "ais_page" selalu
     start dari situ; (2) akses publik ?page=repositori TANPA login, yang
     dieksekusi PALING AWAL bahkan sebelum blok Crawl Berita sempat jalan
     (lihat app.py, akses_publik_repositori) — di jalur ini CSS
     .main-header dari app.py TIDAK PERNAH ke-inject, jadi kartu judul akan
     tampil polos tanpa gaya buat pengunjung publik kalau tidak didefinisikan
     ulang di sini. Nilai persis disamakan dengan app.py biar konsisten.
  */
  .main-header {
    background: linear-gradient(135deg, #0D1B2A 0%, #1C3D5A 100%);
    padding: 14px 20px; border-radius: 10px; margin-bottom: 14px;
    border-bottom: 3px solid #F5A623;
  }
  .main-header h1 { font-size: 1.6rem; font-weight: 700; margin: 0 0 2px 0; color: #F5A623; }
  .main-header p  { font-size: 0.85rem; margin: 0; font-family: 'IBM Plex Mono', monospace; color: rgba(255,255,255,0.75); }

  .repo-card {
    border: 1px solid rgba(245,166,35,0.3);
    border-left: 4px solid #F5A623;
    border-radius: 8px;
    background: rgba(245,166,35,0.05);
    padding: 14px 18px;
    margin-bottom: 10px;
  }
  .repo-tag {
    display: inline-block;
    font-size: 10px;
    font-weight: 600;
    padding: 2px 8px;
    border-radius: 10px;
    background: rgba(245,166,35,0.15);
    color: #F5A623;
    margin-right: 6px;
  }

  /* Kartu sub-analisis — tiap bagian punya warna semantik berbeda */
  .sub-card {
    border-radius: 8px;
    padding: 12px 16px;
    margin-bottom: 10px;
  }
  .sub-label {
    font-size: 10px;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.06em;
    margin-bottom: 6px;
  }
  .sub-text {
    font-size: 13px;
    line-height: 1.6;
  }

  .sub-card.kondisi {
    border-left: 4px solid #94A3B8;
    background: rgba(148,163,184,0.08);
  }
  .sub-card.kondisi .sub-label { color: #94A3B8; }

  .sub-card.dampak {
    border-left: 4px solid #EF4444;
    background: rgba(239,68,68,0.08);
  }
  .sub-card.dampak .sub-label { color: #EF4444; }

  .sub-card.gap {
    border-left: 4px solid #F5A623;
    background: rgba(245,166,35,0.08);
  }
  .sub-card.gap .sub-label { color: #F5A623; }

  .sub-card.usulan {
    border-left: 4px solid #22C55E;
    background: rgba(34,197,94,0.08);
  }
  .sub-card.usulan .sub-label { color: #22C55E; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="main-header">
  <h1>🗄️ Repositori Isu Strategis</h1>
  <p>Arsip hasil yang sudah direview — Pusat Strategi Kebijakan Pengawasan BPKP</p>
</div>
""", unsafe_allow_html=True)


# ── KONEKSI GOOGLE DRIVE ────────────────────────────────────────────────
@st.cache_resource
def get_drive_service():
    """Bangun service object Google Drive dari credentials di Secrets.
    Cache_resource agar koneksi tidak dibangun ulang tiap rerun."""
    from google.oauth2 import service_account
    from googleapiclient.discovery import build

    if not hasattr(st, "secrets") or "gcp_service_account" not in st.secrets:
        return None

    creds_dict = dict(st.secrets["gcp_service_account"])
    credentials = service_account.Credentials.from_service_account_info(
        creds_dict,
        scopes=["https://www.googleapis.com/auth/drive.readonly"],
    )
    return build("drive", "v3", credentials=credentials)


@st.cache_data(ttl=300, show_spinner=False)
def list_excel_files(folder_id: str):
    """List semua file .xlsx di dalam folder Drive tertentu.
    Cache 5 menit — folder ini diisi manual, jadi tidak perlu real-time."""
    service = get_drive_service()
    if service is None:
        return []

    query = (
        f"'{folder_id}' in parents and "
        "mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' and "
        "trashed=false"
    )
    results = service.files().list(
        q=query,
        fields="files(id, name, modifiedTime)",
        orderBy="modifiedTime desc",
    ).execute()
    return results.get("files", [])


@st.cache_data(ttl=300, show_spinner=False)
def download_excel_bytes(file_id: str) -> bytes:
    """Download isi file Excel dari Drive sebagai bytes."""
    from googleapiclient.http import MediaIoBaseDownload

    service = get_drive_service()
    request = service.files().get_media(fileId=file_id)
    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    buf.seek(0)
    return buf.read()


def parse_excel_klaster(file_bytes: bytes, nama_file: str, modified_time: str) -> pd.DataFrame:
    """Parse satu file Excel telaah, kembalikan baris unik per klaster
    yang statusnya Sudah Direview (1 baris representatif per klaster,
    bukan per artikel — karena field telaah identik di semua anggota).
    modified_time adalah tanggal upload/modifikasi file di Google Drive
    (ISO 8601 dari API), dipakai untuk filter kalender di repositori."""
    try:
        df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=0, header=3)
    except Exception:
        return pd.DataFrame()

    kolom_wajib = {"Klaster Isu", "Status Review", "Sektor", "Tema", "Topik"}
    if not kolom_wajib.issubset(set(df.columns)):
        return pd.DataFrame()  # file lama / format tidak kompatibel — skip

    df = df[df["Status Review"] == "Sudah Direview"]
    if len(df) == 0:
        return pd.DataFrame()

    # Satu baris representatif per klaster (ambil baris pertama tiap grup)
    ringkas = df.groupby("Klaster Isu", as_index=False).first()
    ringkas["_sumber_file"] = nama_file
    ringkas["_tanggal_upload"] = pd.to_datetime(modified_time).date()
    return ringkas[[
        "Klaster Isu", "Sektor", "Tema", "Topik",
        "Kondisi/Pemicu Klaster", "Risiko", "Area Perhatian",
        "Dampak/Implikasi (Final)", "Gap Pengawasan", "Usulan Pengawasan",
        "Relevansi Pengawasan", "_sumber_file", "_tanggal_upload",
    ]]


# ── AMBIL & GABUNGKAN SEMUA DATA ─────────────────────────────────────────
folder_id = st.secrets.get("REPOSITORI_FOLDER_ID", "") if hasattr(st, "secrets") else ""

if not folder_id:
    st.warning("⚠️ `REPOSITORI_FOLDER_ID` belum dikonfigurasi di Streamlit Secrets. Hubungi pengelola aplikasi.")
    st.stop()

service_check = get_drive_service()
if service_check is None:
    st.warning("⚠️ Kredensial Google Drive (`gcp_service_account`) belum dikonfigurasi di Streamlit Secrets.")
    st.stop()

with st.spinner("Memuat repositori dari Google Drive..."):
    try:
        files = list_excel_files(folder_id)
    except Exception as e:
        st.error(f"Gagal mengakses folder Google Drive: {str(e)[:300]}")
        st.caption("Pastikan folder sudah di-share ke email service account dengan akses minimal Viewer.")
        st.stop()

    if not files:
        st.info("📭 Folder repositori masih kosong. Upload file Excel hasil telaah yang sudah direview ke folder Google Drive yang sudah ditentukan.")
        st.stop()

    semua_klaster = []
    for f in files:
        try:
            file_bytes = download_excel_bytes(f["id"])
            ringkas = parse_excel_klaster(file_bytes, f["name"], f.get("modifiedTime", ""))
            if len(ringkas):
                semua_klaster.append(ringkas)
        except Exception:
            continue  # file rusak/tidak kompatibel — skip, jangan hentikan seluruh load

    if not semua_klaster:
        st.info("📭 Belum ada klaster dengan status 'Sudah Direview' di file-file yang ada di folder ini.")
        st.stop()

    df_repo = pd.concat(semua_klaster, ignore_index=True)
    df_repo = df_repo.drop_duplicates(subset=["Klaster Isu", "Sektor", "Tema", "Topik"])

st.caption(f"📚 {len(df_repo)} isu dari {len(files)} file di repositori · cache 5 menit")

# ── FILTER TANGGAL UPLOAD ────────────────────────────────────────────────
tgl_min = df_repo["_tanggal_upload"].min()
tgl_max = df_repo["_tanggal_upload"].max()

col_tgl, _ = st.columns([1.4, 2.6])
with col_tgl:
    st.markdown("<div style='font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;opacity:0.55;margin-bottom:4px'>📅 Periode</div>", unsafe_allow_html=True)
    rentang_tanggal = st.date_input(
        "Tanggal Upload", value=(tgl_min, tgl_max),
        label_visibility="collapsed",
    )

# Guard: st.date_input bisa mengembalikan 1 tanggal saja sesaat sebelum
# tanggal akhir dipilih.
if isinstance(rentang_tanggal, tuple) and len(rentang_tanggal) == 2:
    tgl_awal, tgl_akhir = rentang_tanggal
    df_repo = df_repo[
        (df_repo["_tanggal_upload"] >= tgl_awal) &
        (df_repo["_tanggal_upload"] <= tgl_akhir)
    ]
else:
    st.info("Pilih tanggal akhir untuk menerapkan filter rentang.")

# ── Navigasi Sektor — bar chart cakupan ──────────────────────────────────
# Daftar Sektor dari STRUKTUR_APP (bukan nilai unik df_repo) agar sektor
# yang belum ditelaah tetap tampil dengan angka 0.
sektor_counts = df_repo["Sektor"].value_counts().to_dict()
daftar_sektor = list(STRUKTUR_APP.keys())
maks_hitung = max(sektor_counts.values()) if sektor_counts else 1
sektor_aktif = st.session_state.get("repo_sektor_pilih", "Semua Sektor")

st.markdown("<div style='font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;opacity:0.55;margin:16px 0 6px'>🗂️ Sektor</div>", unsafe_allow_html=True)

if sektor_aktif != "Semua Sektor":
    c_label, c_reset = st.columns([5, 1.2])
    with c_label:
        st.markdown(f"<div style='font-size:13px;padding-top:6px'>Menampilkan sektor: <b>{sektor_aktif}</b></div>", unsafe_allow_html=True)
    with c_reset:
        if st.button("✕ Semua Sektor", use_container_width=True):
            st.session_state["repo_sektor_pilih"] = "Semua Sektor"
            st.rerun()

for nama_sektor in daftar_sektor:
    jumlah = int(sektor_counts.get(nama_sektor, 0))
    pct = round(jumlah / maks_hitung * 100) if maks_hitung else 0
    kosong = jumlah == 0
    terpilih = (nama_sektor == sektor_aktif)

    c_bar, c_btn = st.columns([5, 1.2])
    with c_bar:
        st.markdown(f"""
        <div style='display:flex;align-items:center;gap:10px;padding:5px 0'>
          <div style='width:220px;flex-shrink:0;font-size:13px;{"opacity:0.4" if kosong else ""}{"font-weight:600" if terpilih else ""}'>{nama_sektor}</div>
          <div style='flex:1;background:rgba(148,163,184,0.15);border-radius:4px;height:20px;position:relative;overflow:hidden'>
            <div style='width:{pct}%;height:100%;background:{"rgba(148,163,184,0.35)" if kosong else "#F5A623"};border-radius:4px'></div>
          </div>
          <div style='width:26px;text-align:right;font-size:13px;font-weight:600;{"opacity:0.4" if kosong else ""}'>{jumlah}</div>
        </div>
        """, unsafe_allow_html=True)
    with c_btn:
        if not kosong and not terpilih:
            if st.button("Lihat →", key=f"repo_pilih_sektor_{nama_sektor}", use_container_width=True):
                st.session_state["repo_sektor_pilih"] = nama_sektor
                st.rerun()

df_tahap1 = df_repo if sektor_aktif == "Semua Sektor" else df_repo[df_repo["Sektor"] == sektor_aktif]

# ── NAVIGASI TEMA / TOPIK — TETAP DROPDOWN, MENGIKUTI SEKTOR TERPILIH ───
st.markdown("<div style='font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;opacity:0.55;margin:16px 0 6px'>Perincian lebih lanjut</div>", unsafe_allow_html=True)
col_f2, col_f3 = st.columns(2)

with col_f2:
    tema_opsi = ["Semua Tema"] + sorted(df_tahap1["Tema"].dropna().unique().tolist())
    tema_pilih = st.selectbox("Tema", tema_opsi)

df_tahap2 = df_tahap1 if tema_pilih == "Semua Tema" else df_tahap1[df_tahap1["Tema"] == tema_pilih]

with col_f3:
    topik_opsi = ["Semua Topik"] + sorted(df_tahap2["Topik"].dropna().unique().tolist())
    topik_pilih = st.selectbox("Topik", topik_opsi)

df_final = df_tahap2 if topik_pilih == "Semua Topik" else df_tahap2[df_tahap2["Topik"] == topik_pilih]

st.markdown(f"**{len(df_final)} isu** ditemukan sesuai navigasi.")
st.divider()

# ── DAFTAR HASIL ──────────────────────────────────────────────────────────
for _, row in df_final.iterrows():
    with st.expander(f"🗂️ {row['Klaster Isu']}"):
        st.markdown(f"""
        <div class="repo-card">
          <span class="repo-tag">{row['Sektor']}</span>
          <span class="repo-tag">{row['Tema']}</span>
          <span class="repo-tag">{row['Topik']}</span>
        </div>
        """, unsafe_allow_html=True)

        kondisi_safe = html_lib.escape(str(row.get("Kondisi/Pemicu Klaster", "-")))
        dampak_safe = html_lib.escape(str(row.get("Dampak/Implikasi (Final)", "-")))
        gap_safe = html_lib.escape(str(row.get("Gap Pengawasan", "-")))
        usulan_safe = html_lib.escape(str(row.get("Usulan Pengawasan", "-")))

        st.markdown(f"""
        <div class="sub-card kondisi">
          <div class="sub-label">📋 Kondisi / Pemicu</div>
          <div class="sub-text">{kondisi_safe}</div>
        </div>
        <div class="sub-card dampak">
          <div class="sub-label">⚠️ Dampak / Implikasi (Final)</div>
          <div class="sub-text">{dampak_safe}</div>
        </div>
        <div class="sub-card gap">
          <div class="sub-label">🔍 Gap Pengawasan</div>
          <div class="sub-text">{gap_safe}</div>
        </div>
        <div class="sub-card usulan">
          <div class="sub-label">✅ Usulan Pengawasan</div>
          <div class="sub-text">{usulan_safe}</div>
        </div>
        """, unsafe_allow_html=True)

        st.caption(f"Sumber file: {row.get('_sumber_file', '-')} · Diupload: {row.get('_tanggal_upload', '-')}")
