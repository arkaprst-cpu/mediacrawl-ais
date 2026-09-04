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
import re
import html as html_lib
from struktur_app import STRUKTUR_APP

# Ikon per Sektor, dipakai di kartu navigasi Sektor. Nama pada kartu selalu
# nama Sektor asli dari STRUKTUR_APP apa adanya (termasuk kode huruf di
# depannya) — ikon ini cuma penanda visual tambahan, bukan pengganti nama.
SEKTOR_IKON = {
    "A": "🎓", "B": "🤝", "C": "🏗️", "D": "💰",
    "E": "📈", "F": "🌾", "G": "⚡", "H": "🏛️",
}


def _slug(teks: str) -> str:
    """Ubah teks bebas (nama Tema) jadi token aman buat CSS selector / widget key."""
    return re.sub(r"[^a-zA-Z0-9]+", "_", teks).strip("_")[:40] or "x"

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

  /* Kartu navigasi Sektor & Tema — gaya stat-tile (dipinjam dari app.py),
     ukuran seragam untuk semua kartu dalam 1 baris. Sengaja TIDAK dibuat
     proporsional/bento (kartu besar untuk sektor dengan hitungan tinggi)
     karena sebaran data Sektor biasanya sangat timpang (mis. 1 sektor
     terisi, sisanya 0) — ukuran berbeda-beda akan terlihat rusak/tidak
     seimbang dengan sebaran seperti itu. Tidak ada kartu "Semua" lagi di
     grid (supaya jumlah kartu Sektor selalu genap 8 = 2 baris rapi) — reset
     ke "semua" dipindah ke tombol kecil terpisah di sebelah label bagian.
  */
  [class*="st-key-repo_sektor_card_"] button,
  [class*="st-key-repo_tema_card_"] button {
    background: rgba(128,128,128,0.06) !important;
    border: 1px solid rgba(128,128,128,0.2) !important;
    border-radius: 10px !important;
    height: auto !important;
    line-height: 1.4 !important;
    transition: background 0.15s ease, border-color 0.15s ease;
  }
  [class*="st-key-repo_sektor_card_"] button {
    padding: 1.3rem 1rem !important;
  }
  [class*="st-key-repo_sektor_card_"] button p {
    font-size: 0.85rem !important;
  }
  [class*="st-key-repo_tema_card_"] button p {
    font-size: 0.78rem !important;
  }
  [class*="st-key-repo_sektor_card_"] button:hover:not(:disabled),
  [class*="st-key-repo_tema_card_"] button:hover:not(:disabled) {
    background: rgba(245,166,35,0.1) !important;
    border-color: rgba(245,166,35,0.4) !important;
  }
  /* Tombol reset kecil (bukan kartu) di sebelah label Sektor/Tema */
  .st-key-repo_sektor_reset button, .st-key-repo_tema_reset button {
    font-size: 0.72rem !important;
    padding: 0.2rem 0.5rem !important;
    height: auto !important;
    opacity: 0.75;
  }
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

    # Nama kolom ini berubah di buat_excel() (app.py) pada 2026-08-31:
    # "Kondisi/Pemicu Klaster" -> "Kondisi Klaster". Folder Drive berisi
    # file dari SEBELUM dan SESUDAH perubahan itu, jadi dua-duanya harus
    # diterima. Diseragamkan ke satu nama internal di sini supaya kode
    # tampilan di bawah (row.get("Kondisi/Pemicu Klaster")) tidak perlu
    # ikut bercabang.
    if "Kondisi/Pemicu Klaster" not in df.columns:
        if "Kondisi Klaster" in df.columns:
            df = df.rename(columns={"Kondisi Klaster": "Kondisi/Pemicu Klaster"})
        else:
            df = df.assign(**{"Kondisi/Pemicu Klaster": "-"})

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

# ── PENCARIAN CEPAT — jalan pintas, melewati hierarki Sektor→Tema→Topik ──
st.markdown("<div style='font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;opacity:0.55;margin:16px 0 6px'>🔎 Cari Cepat</div>", unsafe_allow_html=True)

col_search, col_clear = st.columns([5, 1.2])
# Tombol "Hapus" di-render lebih dulu di urutan kode (meski tampil di kanan
# lewat kolom) supaya perubahan session_state["repo_cari_teks"]-nya terjadi
# SEBELUM widget text_input dengan key yang sama diinstansiasi di bawah —
# Streamlit melarang mengubah state widget setelah ia dibuat di run yang sama.
with col_clear:
    if st.session_state.get("repo_cari_teks", "").strip():
        st.markdown("<div style='height:1.6rem'></div>", unsafe_allow_html=True)
        if st.button("✕ Hapus", key="repo_cari_clear_btn", use_container_width=True):
            st.session_state["repo_cari_teks"] = ""
            st.rerun()
with col_search:
    cari_teks = st.text_input(
        "Cari klaster isu", placeholder="Cari klaster isu... (melewati navigasi Sektor/Tema/Topik)",
        label_visibility="collapsed", key="repo_cari_teks",
    )

mode_cari = bool(cari_teks.strip())

if mode_cari:
    df_final = df_repo[df_repo["Klaster Isu"].str.contains(cari_teks.strip(), case=False, na=False, regex=False)]
    st.markdown(f"**{len(df_final)} isu** cocok dengan pencarian \"{cari_teks.strip()}\".")

else:
    # ── Navigasi Sektor — kartu grid ala stat-tile ──────────────────────
    # Daftar Sektor dari STRUKTUR_APP (bukan nilai unik df_repo) agar sektor
    # yang belum ditelaah tetap tampil dengan angka 0.
    sektor_counts = df_repo["Sektor"].value_counts().to_dict()
    daftar_sektor = list(STRUKTUR_APP.keys())
    sektor_aktif = st.session_state.get("repo_sektor_pilih", "Semua Sektor")

    # Reset navigasi Tema begitu Sektor berubah, supaya tidak nyangkut di
    # Tema yang sudah tidak relevan dengan Sektor yang baru dipilih.
    if st.session_state.get("_repo_sektor_prev") != sektor_aktif:
        st.session_state["repo_tema_pilih"] = "Semua Tema"
        st.session_state["_repo_sektor_prev"] = sektor_aktif

    # "" (tidak ada kartu aktif) waktu di Semua Sektor — tidak ada kartu
    # "Semua" lagi di grid, resetnya lewat tombol kecil di sebelah label.
    sektor_kode_aktif = "" if sektor_aktif == "Semua Sektor" else sektor_aktif.split(".")[0].strip()

    st.markdown(f"""
    <style>
    .st-key-repo_sektor_card_{sektor_kode_aktif} button {{
        background: rgba(245,166,35,0.16) !important;
        border: 1px solid rgba(245,166,35,0.55) !important;
    }}
    .st-key-repo_sektor_card_{sektor_kode_aktif} button p {{ color: #F5A623 !important; font-weight: 700 !important; }}
    </style>
    """, unsafe_allow_html=True)

    c_label1, c_reset1 = st.columns([5, 1.6])
    with c_label1:
        st.markdown("<div style='font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;opacity:0.55;margin:16px 0 8px'>🗂️ Sektor</div>", unsafe_allow_html=True)
    with c_reset1:
        if sektor_aktif != "Semua Sektor":
            if st.button("↺ Semua Sektor", key="repo_sektor_reset", use_container_width=True):
                st.session_state["repo_sektor_pilih"] = "Semua Sektor"
                st.rerun()

    # Nama kartu SELALU nama Sektor asli dari STRUKTUR_APP apa adanya
    # (termasuk kode huruf A/B/C/... di depannya) — tidak diparafrase.
    kartu_sektor = [
        {"kode": nama_sektor.split(".")[0].strip(), "nama": nama_sektor,
         "jumlah": int(sektor_counts.get(nama_sektor, 0))}
        for nama_sektor in daftar_sektor
    ]

    for baris_awal in range(0, len(kartu_sektor), 4):
        kolom = st.columns(4)
        for i, kartu in enumerate(kartu_sektor[baris_awal:baris_awal + 4]):
            with kolom[i]:
                with st.container(key=f"repo_sektor_card_{kartu['kode']}"):
                    kosong = kartu["jumlah"] == 0
                    icon = SEKTOR_IKON.get(kartu["kode"], "📌")
                    label = f"{icon}\n\n**{kartu['jumlah']}**\n\n{kartu['nama']}"
                    if st.button(label, key=f"repo_sektor_btn_{kartu['kode']}", use_container_width=True,
                                 disabled=kosong, help=kartu["nama"]):
                        if kartu["nama"] != sektor_aktif:
                            st.session_state["repo_sektor_pilih"] = kartu["nama"]
                            st.rerun()

    df_tahap1 = df_repo if sektor_aktif == "Semua Sektor" else df_repo[df_repo["Sektor"] == sektor_aktif]

    # ── Navigasi Tema — kartu/chip, mengikuti Sektor terpilih ───────────
    tema_counts = df_tahap1["Tema"].value_counts().to_dict()
    daftar_tema = sorted(df_tahap1["Tema"].dropna().unique().tolist())
    tema_aktif = st.session_state.get("repo_tema_pilih", "Semua Tema")
    if tema_aktif != "Semua Tema" and tema_aktif not in daftar_tema:
        tema_aktif = "Semua Tema"
        st.session_state["repo_tema_pilih"] = "Semua Tema"

    tema_kode_aktif = "" if tema_aktif == "Semua Tema" else _slug(tema_aktif)

    st.markdown(f"""
    <style>
    .st-key-repo_tema_card_{tema_kode_aktif} button {{
        background: rgba(245,166,35,0.16) !important;
        border: 1px solid rgba(245,166,35,0.55) !important;
    }}
    .st-key-repo_tema_card_{tema_kode_aktif} button p {{ color: #F5A623 !important; font-weight: 700 !important; }}
    </style>
    """, unsafe_allow_html=True)

    c_label2, c_reset2 = st.columns([5, 1.6])
    with c_label2:
        st.markdown("<div style='font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;opacity:0.55;margin:16px 0 8px'>🏷️ Tema</div>", unsafe_allow_html=True)
    with c_reset2:
        if tema_aktif != "Semua Tema":
            if st.button("↺ Semua Tema", key="repo_tema_reset", use_container_width=True):
                st.session_state["repo_tema_pilih"] = "Semua Tema"
                st.rerun()

    kartu_tema = [{"kode": _slug(t), "nama": t, "jumlah": int(tema_counts.get(t, 0))} for t in daftar_tema]

    if kartu_tema:
        for baris_awal in range(0, len(kartu_tema), 3):
            kolom = st.columns(3)
            for i, kartu in enumerate(kartu_tema[baris_awal:baris_awal + 3]):
                with kolom[i]:
                    with st.container(key=f"repo_tema_card_{kartu['kode']}"):
                        label_pendek = kartu["nama"] if len(kartu["nama"]) <= 30 else kartu["nama"][:28] + "…"
                        label = f"{label_pendek}  ·  {kartu['jumlah']}"
                        if st.button(label, key=f"repo_tema_btn_{kartu['kode']}", use_container_width=True, help=kartu["nama"]):
                            if kartu["nama"] != tema_aktif:
                                st.session_state["repo_tema_pilih"] = kartu["nama"]
                                st.rerun()
    else:
        st.caption("Tidak ada Tema untuk Sektor ini.")

    df_tahap2 = df_tahap1 if tema_aktif == "Semua Tema" else df_tahap1[df_tahap1["Tema"] == tema_aktif]

    # ── Topik — tetap dropdown ───────────────────────────────────────────
    st.markdown("<div style='font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;opacity:0.55;margin:16px 0 6px'>Topik</div>", unsafe_allow_html=True)
    topik_opsi = ["Semua Topik"] + sorted(df_tahap2["Topik"].dropna().unique().tolist())
    topik_pilih = st.selectbox("Topik", topik_opsi, label_visibility="collapsed")

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
