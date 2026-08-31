"""
dashboard_ais.py
Halaman Dashboard AIS — Analisis Isu Strategis Pengawasan
Pusat Strategi Kebijakan Pengawasan BPKP

Sumber data: hasil crawl sesi aktif (session_state) atau upload Excel (.xlsx)
via sidebar.
"""

import streamlit as st
import pandas as pd
import json
import io
import re
from collections import Counter
from datetime import datetime
from struktur_app import STRUKTUR_APP


def pisahkan_sumber_judul(judul: str):
    """Pisahkan judul dari nama sumber yang ditempel Google News di akhir
    (format '... - NamaSumber'). Mengembalikan (judul_bersih, nama_sumber)."""
    s = str(judul)
    m = re.search(r"\s[-–]\s([^-–]+)$", s)
    if m:
        return s[:m.start()].strip(), m.group(1).strip()
    return s.strip(), ""


def pill_sumber_html(sumber: str, compact: bool = False) -> str:
    """Render pill kecil nama sumber (Kompas, CNN, Tempo, dst.) di awal judul."""
    if not sumber:
        return ""
    cls = "pill-sumber pill-sumber-compact" if compact else "pill-sumber"
    return f'<span class="{cls}">{sumber}</span>'


# ── CUSTOM CSS ───────────────────────────────────────────────
st.markdown("""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;600&display=swap');

  html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

  /* Panel kanan fixed via st.container(key="panel_kanan") — widget
     interaktif Streamlit tidak bisa dibungkus position:fixed lewat HTML
     markdown biasa. Kalibrasi ini (top:90px) sempat diganti dua kali
     (sticky, lalu fixed-dari-bawah) untuk mengejar bug "nabrak topbar"
     di Streamlit Cloud, tapi dua-duanya bikin regresi lebih parah:
     sticky bikin panel hilang total saat scroll panjang, dan
     fixed-dari-bawah memaksa max-height sangat pendek yang
     mengempeskan tinggi field form Telaah Klaster (text area jadi
     cuma garis tipis). Sengaja dikembalikan ke versi ini — lebih
     diterima sesekali nabrak topbar (kosmetik) daripada form telaah
     tidak bisa dipakai (fungsional). max-height dilebarkan hampir
     sepenuh tinggi viewport (cuma sisa 16px margin bawah) supaya
     ruang mengetik uraian panjang di form Telaah Klaster lebih lega. */
  .st-key-panel_kanan {
    position: fixed !important;
    top: 90px !important;
    right: 24px !important;
    width: min(42vw, 520px) !important;
    max-height: calc(100vh - 106px) !important;
    overflow-y: auto !important;
    z-index: 999 !important;
    background: rgba(13,27,42,0.97) !important;
    border: 1px solid rgba(245,166,35,0.35) !important;
    border-top: 4px solid #F5A623 !important;
    border-radius: 10px !important;
    padding: 16px 18px !important;
    box-shadow: 0 8px 32px rgba(0,0,0,0.4) !important;
  }

  /* Kartu upload di tengah landing state (belum ada data) */
  .st-key-landing_upload_card {
    border: 2px dashed rgba(128,128,128,0.3) !important;
    border-radius: 12px !important;
    background: rgba(128,128,128,0.06) !important;
    padding: 24px 24px 16px !important;
  }

  /* Topbar */
  .ais-topbar {
    background: linear-gradient(135deg, #0D1B2A 0%, #1C3D5A 100%);
    border-radius: 8px;
    padding: 16px 24px;
    margin-bottom: 20px;
    border-bottom: 3px solid #F5A623;
    display: flex;
    align-items: center;
    justify-content: space-between;
  }
  .ais-logo { font-family: 'JetBrains Mono', monospace; font-size: 22px; font-weight: 700; color: #F5A623; }
  .ais-subtitle { font-size: 12px; color: rgba(255,255,255,0.65); margin-top: 2px; }
  .ais-badge {
    background: #F5A623; color: #0D1B2A;
    font-family: 'JetBrains Mono', monospace;
    font-size: 11px; font-weight: 700;
    padding: 4px 12px; border-radius: 4px;
    letter-spacing: 0.06em;
  }

  /* Spotlight */
  .spotlight-box {
    background: linear-gradient(135deg, #0D1B2A 0%, #1C3D5A 100%);
    border-radius: 8px;
    padding: 20px 24px;
    margin-bottom: 16px;
    border-left: 4px solid #F5A623;
  }
  .spotlight-eyebrow {
    font-size: 10px; font-weight: 700;
    letter-spacing: 0.12em; text-transform: uppercase;
    color: #F5A623; margin-bottom: 6px;
  }
  .spotlight-title { font-size: 16px; font-weight: 700; color: white; line-height: 1.3; margin-bottom: 10px; }
  .spotlight-body { font-size: 12px; color: rgba(255,255,255,0.75); line-height: 1.65; }

  /* Stat cards — dua tingkat: primary (metrik paling actionable, lebih
     besar) dan secondary (rincian pendukung, lebih kecil) supaya tidak
     semua angka bersaing dengan bobot yang sama. */
  .stat-card-primary {
    background: rgba(128,128,128,0.06); border: 1px solid rgba(128,128,128,0.2);
    border-radius: 8px; padding: 20px;
    text-align: center;
    box-shadow: 0 2px 10px rgba(0,0,0,0.08);
  }
  .stat-num-primary {
    font-family: 'JetBrains Mono', monospace;
    font-size: 38px; font-weight: 700; color: inherit; line-height: 1;
  }
  .stat-label-primary {
    font-size: 12px; color: inherit; opacity: 0.65; margin-top: 6px;
    text-transform: uppercase; letter-spacing: 0.04em;
  }
  .stat-card-secondary {
    background: rgba(128,128,128,0.04); border: 1px solid rgba(128,128,128,0.15);
    border-radius: 6px; padding: 10px;
    text-align: center;
  }
  .stat-num-secondary {
    font-family: 'JetBrains Mono', monospace;
    font-size: 20px; font-weight: 700; color: inherit; line-height: 1;
  }
  .stat-label-secondary { font-size: 10px; color: inherit; opacity: 0.55; margin-top: 3px; }

  /* Issue card — base dipakai apa adanya untuk daftar datar (fallback
     tanpa klaster). Dua varian di bawah menyesuaikan bobot visual sesuai
     konteks: -highlight untuk sorotan "perlu perhatian" (Tab Ikhtisar),
     -member untuk anggota klaster yang harus terasa lebih ringan daripada
     blok INDUK KLASTER di atasnya (Tab Klasterisasi Isu). */
  .issue-card {
    background: rgba(128,128,128,0.06); border: 1px solid rgba(128,128,128,0.2);
    border-radius: 6px; padding: 14px 14px 14px 18px;
    margin-bottom: 8px; position: relative;
    overflow: hidden;
  }
  .issue-card::before {
    content: ''; position: absolute;
    left: 0; top: 0; bottom: 0; width: 4px;
  }
  .issue-card.negatif::before { background: #E74C3C; }
  .issue-card.netral::before { background: #95A5A6; }
  .issue-card.positif::before { background: #27AE60; }
  .issue-title { font-size: 13px; font-weight: 600; color: inherit; line-height: 1.4; }
  .issue-sub { font-size: 11px; color: inherit; opacity: 0.65; margin-top: 2px; }
  .issue-summary { font-size: 11px; color: inherit; opacity: 0.65; line-height: 1.5; margin-top: 6px; }

  .issue-card-highlight {
    padding: 16px 16px 16px 20px;
    box-shadow: 0 2px 10px rgba(0,0,0,0.12);
  }
  .issue-card-highlight::before { width: 5px; }
  .issue-card-highlight .issue-title { font-size: 14px; font-weight: 700; }

  .issue-card-member {
    padding: 9px 12px 9px 14px;
    background: transparent;
  }
  .issue-card-member::before { width: 3px; }
  .issue-card-member .issue-title { font-size: 12px; font-weight: 500; opacity: 0.85; }

  /* Kartu artikel + tombol "Lihat detail →" — dibungkus satu
     st.container(key="artikel_card_N") di Python supaya keduanya satu
     elemen DOM, lalu di sini disambungkan jadi satu unit visual: card
     persegi di atas (radius bawah dihilangkan), tombol jadi strip footer
     tanpa celah/border ganda di bawahnya. Sebelumnya tombolnya render
     sebagai widget terpisah di bawah card HTML, jadi ambigu itu tombol
     punya card yang mana. */
  [class*="st-key-artikel_card_"] [data-testid="stVerticalBlock"] { gap: 0 !important; }
  [class*="st-key-artikel_card_"] .issue-card.artikel-card-attached {
    border-radius: 6px 6px 0 0;
    margin-bottom: 0;
  }
  [class*="st-key-artikel_card_"] div[data-testid="stButton"] { margin: 0; }
  [class*="st-key-artikel_card_"] div[data-testid="stButton"] button {
    width: 100%;
    border: 1px solid rgba(128,128,128,0.2);
    border-top: none;
    border-radius: 0 0 6px 6px;
    background: rgba(128,128,128,0.04);
    color: inherit; opacity: 0.7;
    font-size: 11px; font-weight: 500;
    padding: 5px 12px; min-height: 30px;
  }
  [class*="st-key-artikel_card_"] div[data-testid="stButton"] button:hover {
    background: rgba(99,179,237,0.14);
    border-color: rgba(99,179,237,0.45);
    color: #63B3ED; opacity: 1;
  }

  /* Pill sumber — nama media (Kompas, CNN, Tempo, dst.) ditonjolkan di
     awal judul supaya pembaca langsung tahu asal beritanya. */
  .pill-sumber {
    display: inline-block; font-size: 10px; font-weight: 700;
    font-family: 'JetBrains Mono', monospace;
    color: #F5A623; background: rgba(245,166,35,0.12);
    border: 1px solid rgba(245,166,35,0.35);
    padding: 1px 7px; border-radius: 3px;
    margin-right: 6px; letter-spacing: 0.02em;
    text-transform: uppercase; vertical-align: middle;
  }
  .spotlight-title .pill-sumber { vertical-align: 2px; }
  .pill-sumber-compact {
    font-size: 9px; padding: 0px 5px; opacity: 0.75;
  }

  /* Badges */
  .badge {
    display: inline-block; font-size: 10px; font-weight: 600;
    font-family: 'JetBrains Mono', monospace;
    padding: 2px 7px; border-radius: 3px;
    margin-right: 4px; margin-top: 2px;
  }
  .badge-negatif { background: rgba(231,76,60,0.15); color: #E74C3C; }
  .badge-netral { background: rgba(127,140,141,0.15); color: #95A5A6; }
  .badge-positif { background: rgba(39,174,96,0.15); color: #27AE60; }
  /* Topik/subisu (kategori isu) vs Aktor/Lokasi (pihak yang terlibat)
     tadinya sama-sama pakai .badge-aktor sehingga tidak terlihat beda
     level informasinya. Dipisah: -topik tetap indigo (kategori),
     -aktor jadi slate + ikon 👤 (entitas/pihak). */
  .badge-topik { background: rgba(99,102,241,0.15); color: #818CF8; }
  /* Tag dimensi pengawasan (Governance/Risk/Control/Compliance/
     Anti-Korupsi/Debottlenecking) di kartu INDUK KLASTER — warna
     senada aksen oranye klaster (#F5A623) tapi lebih redup, supaya
     kebaca sebagai "klasifikasi ringkas", bukan bersaing dengan judul. */
  .badge-dimensi {
    background: rgba(245,166,35,0.14); color: #F5A623;
    border: 1px solid rgba(245,166,35,0.3);
  }
  .badge-aktor { background: rgba(148,163,184,0.16); color: #94A3B8; }

  /* Detail box */
  .detail-section { margin-bottom: 14px; }
  .detail-label {
    font-size: 10px; font-weight: 700;
    text-transform: uppercase; letter-spacing: 0.08em;
    color: inherit; opacity: 0.55; margin-bottom: 4px;
  }
  .detail-text { font-size: 12px; color: inherit; line-height: 1.6; }
  .link-artikel-asli {
    color: #5AA9FF; font-weight: 600; font-size: 11px;
    text-decoration: underline; text-underline-offset: 2px;
  }
  .link-artikel-asli:hover { color: #8CC4FF; }

  /* Expander klaster (Tab Klasterisasi Isu) — beri identitas visual "kartu
     klaster" pada header accordion-nya sendiri, bukan cuma pada isinya
     saat dibuka. Tanpa ini header-nya kelihatan sama dengan accordion
     generik, padahal secara fungsi dia adalah unit pengelompokan
     artikel (klaster), bukan baris daftar biasa. */
  [data-testid="stExpander"] {
    border: none !important;
    margin-bottom: 12px !important;
  }
  [data-testid="stExpander"] > details {
    border: 1px solid rgba(245,166,35,0.4) !important;
    border-left: 5px solid #F5A623 !important;
    border-radius: 8px !important;
    background: rgba(245,166,35,0.07) !important;
    overflow: hidden !important;
  }
  [data-testid="stExpander"] summary {
    padding: 14px 16px !important;
  }
  [data-testid="stExpander"] summary [data-testid="stMarkdownContainer"] p {
    font-size: 14px !important;
    font-weight: 700 !important;
    letter-spacing: 0.01em;
  }
  [data-testid="stExpander"] summary [data-testid="stIconMaterial"] {
    color: #F5A623 !important;
  }
  [data-testid="stExpanderDetails"] {
    background: rgba(13,27,42,0.35) !important;
    border-top: 1px solid rgba(245,166,35,0.15) !important;
  }
  .implikasi-box {
    background: rgba(245,166,35,0.12); border: 1px solid rgba(245,166,35,0.5);
    border-radius: 5px; padding: 10px 12px;
    font-size: 12px; color: inherit; line-height: 1.6;
  }
  .tindaklanjut-box {
    background: rgba(147,180,232,0.12); border: 1px solid rgba(147,180,232,0.5);
    border-radius: 5px; padding: 10px 12px;
    font-size: 12px; color: inherit; line-height: 1.6;
  }

  /* Tone pill */
  .tone-neg { color: #E74C3C; font-weight: 600; }
  .tone-net { color: #7F8C8D; font-weight: 600; }
  .tone-pos { color: #27AE60; font-weight: 600; }

  /* Hide streamlit chrome */
  #MainMenu {visibility: hidden;}
  footer {visibility: hidden;}
  .block-container { padding-top: 1rem; padding-bottom: 1rem; }

  /* Uploader di sidebar — dropzone bawaan Streamlit besar padahal cuma
     dipakai sesekali per sesi; diperkecil agar tidak mendominasi sidebar. */
  [data-testid="stFileUploaderDropzone"] {
    padding: 8px 12px !important;
    min-height: 0 !important;
  }
  [data-testid="stFileUploaderDropzoneInstructions"] svg { display: none; }
  [data-testid="stFileUploaderDropzoneInstructions"] span { font-size: 11px !important; }
  [data-testid="stFileUploaderDropzoneInstructions"] small { display: none; }
  [data-testid="stBaseButton-secondary"] { padding: 2px 10px !important; font-size: 11px !important; }
</style>
""", unsafe_allow_html=True)


# ── DATA LOADER ──────────────────────────────────────────────
def load_from_excel(uploaded_file):
    """Parse Excel output dari pipeline AIS. Mendukung dua format:
    - 12 kolom (dengan Klaster Isu, tanpa Kondisi/Pemicu & Relevansi)
    - 14 kolom (lengkap, dengan Kondisi/Pemicu Klaster & Relevansi Pengawasan)
    Format lama 11 kolom (tanpa Klaster) tidak lagi didukung.
    """
    df_raw = pd.read_excel(uploaded_file, sheet_name=0, header=None)

    # Baca metadata dari baris ke-2 (index 1)
    meta_str = str(df_raw.iloc[1, 0]) if df_raw.shape[0] > 1 else ""
    meta = {"raw": meta_str}
    try:
        parts = meta_str.split("|")
        meta["isu"] = parts[0].replace("Isu:", "").strip() if len(parts) > 0 else "—"
        meta["generate"] = parts[1].replace("Generate:", "").strip() if len(parts) > 1 else "—"
        meta["total"] = parts[2].replace("Total:", "").replace("artikel", "").strip() if len(parts) > 2 else "—"
        meta["unit"] = parts[3].strip() if len(parts) > 3 else "Pusat Strategi Kebijakan Pengawasan BPKP"
    except:
        meta["isu"] = "BPKP"; meta["generate"] = "—"; meta["total"] = "—"; meta["unit"] = "Pusat Strategi Kebijakan Pengawasan BPKP"

    # Data mulai dari baris ke-4 (header di index 2, data mulai index 3)
    df = pd.read_excel(uploaded_file, sheet_name=0, header=3)
    n_kolom = df.shape[1]

    kolom_telaah = ['Sektor','Tema','Topik','DampakImplikasiFinal','GapPengawasan','UsulanPengawasan','StatusReview']
    KOLOM_INTI_14 = ['No','Klaster','Tanggal','Sumber','Link','Judul','Ringkasan','IsuSubisu','AktorLokasi','Tone','Risiko','TindakLanjut','KondisiPemicu','RelevansiPengawasan']

    # PENTING — jenjang format ditambah dari BAWAH (n_kolom>=22), bukan
    # mengubah cabang lama: kolom 1-21 (termasuk blok telaah manusia
    # Sektor..Status Review di 15-21) TIDAK PERNAH digeser posisinya. File
    # lama (21 kolom, dari sebelum kolom Dimensi Pengawasan ditambahkan)
    # tetap lewat cabang n_kolom>=21 apa adanya — status "Sudah Direview"
    # tidak boleh ikut hilang/reset hanya karena kolom baru ditambahkan.
    if n_kolom >= 22:
        # Format terbaru (22 kolom): 14 kolom inti + 7 kolom telaah manusia
        # + 1 kolom Dimensi Pengawasan (GRCC AnCoDe) di paling akhir.
        df_dimensi = df.iloc[:, 21:22].copy()
        df_dimensi.columns = ['DimensiPengawasan']
        df_telaah = df.iloc[:, 14:21].copy()
        df_telaah.columns = kolom_telaah
        df = df.iloc[:, :14]
        df.columns = KOLOM_INTI_14
        df = pd.concat([df, df_telaah, df_dimensi], axis=1)
    elif n_kolom >= 21:
        # Format lengkap hasil telaah (21 kolom, TANPA Dimensi Pengawasan —
        # file dari sebelum fitur itu ada). Kolom 15-21 berisi hasil Human
        # Review (Sektor..Status Review). Wajib dipertahankan supaya status
        # "Sudah Direview" tidak hilang saat file ini di-upload lagi untuk
        # melanjutkan kerja telaah.
        df_telaah = df.iloc[:, 14:21].copy()
        df_telaah.columns = kolom_telaah
        df = df.iloc[:, :14]
        df.columns = KOLOM_INTI_14
        df = pd.concat([df, df_telaah], axis=1)
        df['DimensiPengawasan'] = ''
    elif n_kolom >= 14:
        df = df.iloc[:, :14]
        df.columns = KOLOM_INTI_14
        for k in kolom_telaah:
            df[k] = 'Belum Direview' if k == 'StatusReview' else '-'
        df['DimensiPengawasan'] = ''
    else:
        df.columns = ['No','Klaster','Tanggal','Sumber','Link','Judul','Ringkasan','IsuSubisu','AktorLokasi','Tone','Risiko','TindakLanjut']
        df['KondisiPemicu'] = '-'
        df['RelevansiPengawasan'] = '-'
        for k in kolom_telaah:
            df[k] = 'Belum Direview' if k == 'StatusReview' else '-'
        df['DimensiPengawasan'] = ''

    df = df.dropna(subset=['Judul'])
    df = df[df['No'] != 'No']
    df = df.reset_index(drop=True)
    df['Tanggal'] = df['Tanggal'].astype(str)

    return df, meta


def compute_stats(df):
    """Hitung statistik dari dataframe."""
    tone_counts = df['Tone'].value_counts().to_dict()
    total = len(df)
    neg = tone_counts.get('Negatif', 0)
    net = tone_counts.get('Netral', 0)
    pos = tone_counts.get('Positif', 0)
    
    # Hitung skor risiko sederhana berdasarkan tone + panjang teks risiko
    # Level risiko berbasis tone + panjang teks risiko
    def skor_risiko(row):
        base = {'Negatif': 7, 'Netral': 5, 'Positif': 3}.get(str(row['Tone']), 5)
        bonus = min(2, len(str(row['Risiko'])) // 150)
        return base + bonus

    df = df.copy()
    df['skor_risiko'] = df.apply(skor_risiko, axis=1)
    df['level_risiko'] = df['skor_risiko'].apply(
        lambda s: 'Tinggi' if s >= 8 else ('Sedang' if s >= 6 else 'Rendah')
    )
    
    return df, {
        'total': total,
        'negatif': neg, 'netral': net, 'positif': pos,
        'pct_neg': round(neg/total*100) if total else 0,
        'pct_net': round(net/total*100) if total else 0,
        'pct_pos': round(pos/total*100) if total else 0,
        'tinggi': (df['level_risiko']=='Tinggi').sum(),
        'sedang': (df['level_risiko']=='Sedang').sum(),
    }


def extract_keywords(df):
    """Extract top keywords dari kolom IsuSubisu dan Judul."""
    stopwords = {'dan','di','ke','dari','untuk','yang','ini','itu','dengan',
                 'pada','oleh','sebagai','dalam','telah','akan','dapat','tidak',
                 'bpjs','kesehatan','bpkp','atas','terkait','bagi','juga','serta'}
    words = []
    for col in ['IsuSubisu', 'Judul']:
        for text in df[col].astype(str):
            for w in text.lower().split():
                w = w.strip('.,;:!?()[]"\'')
                if len(w) > 3 and w not in stopwords:
                    words.append(w)
    return Counter(words).most_common(15)


def extract_aktors(df):
    """Extract aktor yang paling sering disebut."""
    aktors = []
    for text in df['AktorLokasi'].astype(str):
        for a in text.split(','):
            a = a.strip()
            if a and a != 'nan' and len(a) > 2:
                aktors.append(a)
    return Counter(aktors).most_common(12)


def risiko_per_aktor(df, top_n=8):
    """Sebaran level risiko per Aktor/Lokasi — siapa yang paling sering
    terkait artikel risiko Tinggi, bukan sekadar paling sering disebut.
    Satu artikel dengan beberapa aktor (dipisah koma) dihitung untuk
    masing-masing aktor."""
    rows = []
    for _, r in df.iterrows():
        aktor_text = str(r['AktorLokasi'])
        if aktor_text == 'nan' or not aktor_text.strip():
            continue
        for a in aktor_text.split(','):
            a = a.strip()
            if a and len(a) > 2:
                rows.append({'aktor': a, 'level_risiko': r['level_risiko']})

    if not rows:
        return []

    df_aktor = pd.DataFrame(rows)
    pivot = df_aktor.groupby(['aktor', 'level_risiko']).size().unstack(fill_value=0)
    for lvl in ['Tinggi', 'Sedang', 'Rendah']:
        if lvl not in pivot.columns:
            pivot[lvl] = 0
    pivot['total'] = pivot[['Tinggi', 'Sedang', 'Rendah']].sum(axis=1)

    # Urutkan: jumlah Tinggi dulu (desc), lalu total (desc) sebagai tie-breaker
    pivot = pivot.sort_values(['Tinggi', 'total'], ascending=[False, False]).head(top_n)

    return [
        {'aktor': idx, 'tinggi': row['Tinggi'], 'sedang': row['Sedang'],
         'rendah': row['Rendah'], 'total': row['total']}
        for idx, row in pivot.iterrows()
    ]


# ── SIDEBAR ──────────────────────────────────────────────────
# Tombol upload sengaja TIDAK selalu di sidebar. Selama belum ada data
# sama sekali (landing state), dia ditaruh di tengah halaman utama —
# itu satu-satunya aksi yang relevan buat user di titik itu, jadi lebih
# masuk akal ditonjolkan di tengah alur baca daripada "disembunyikan"
# di panel kiri yang gampang kelewat. Begitu ada data (dari upload atau
# dari sesi crawl), tempatnya pindah ke sidebar sebagai cara ganti file
# tanpa mengganggu tampilan dashboard.
#
# "has_upload" TIDAK dicek dari nilai widget file_uploader itu sendiri
# (uploaded is not None) — itu penyebab bug "balik ke sidebar lagi"
# yang sempat dilaporkan: app.py ini multipage lewat exec(), jadi saat
# user pindah ke halaman lain, dashboard_ais.py sama sekali tidak
# dieksekusi pada run itu, widget-nya jadi tidak ter-render, dan
# Streamlit MEMBUANG file yang sudah diupload ke widget itu (beda dari
# entri session_state biasa yang tahan pindah halaman). Begitu balik ke
# Dashboard, `uploaded` sudah None lagi walau user merasa baru saja
# upload. Solusinya: begitu file diupload, langsung di-parse dan hasil
# parsingnya (bukan objek file mentahnya) disimpan ke session_state
# ("_dash_upload_df_raw" dkk) — itu yang jadi sumber kebenaran, persis
# seperti "has_session" untuk data dari sesi crawl, dan sama-sama tahan
# pindah halaman.
has_session = st.session_state.get("ais_ready", False) and "hasil" in st.session_state
has_upload = "_dash_upload_df_raw" in st.session_state
uploader_di_sidebar = has_session or has_upload

with st.sidebar:
    if uploader_di_sidebar:
        st.markdown("### 📂 Upload Data")
        st.markdown("<div style='font-size:11px;color:#888;margin-bottom:8px'>Upload file Excel hasil pipeline crawl AIS (.xlsx)</div>", unsafe_allow_html=True)
        uploaded = st.file_uploader(
            "Pilih file Excel", type=["xlsx"],
            label_visibility="collapsed", key="uploader_xlsx"
        )
    else:
        uploaded = None

    # Semua tone & level risiko ditampilkan (tidak ada filter sidebar aktif)
    filter_tone = ["Negatif", "Netral", "Positif"]
    filter_risiko = ["Tinggi", "Sedang", "Rendah"]

    st.markdown("---")
    st.markdown("<div style='font-size:10px;color:#aaa;line-height:1.6'>Analisis Isu Strategis Pengawasan<br>Pusat Strategi Kebijakan Pengawasan BPKP</div>", unsafe_allow_html=True)

# File baru dari uploader sidebar (mis. user ganti file) langsung
# di-parse & disimpan ke session_state — lihat catatan panjang di atas.
if uploaded is not None:
    df_raw_baru, meta_baru = load_from_excel(uploaded)
    st.session_state["_dash_upload_df_raw"] = df_raw_baru
    st.session_state["_dash_upload_meta"] = meta_baru
    st.session_state["_dash_upload_file_id"] = getattr(uploaded, "file_id", None) or f"{uploaded.name}_{uploaded.size}"
    st.session_state["_dash_last_source"] = "upload"
    has_upload = True


# ── MAIN CONTENT ─────────────────────────────────────────────
# Prioritas sumber data — BUKAN lagi "upload selalu menang" secara
# statis. Itu penyebab bug yang dilaporkan user: habis crawl baru (mis.
# 20 artikel), Dashboard AIS malah masih menampilkan Excel upload manual
# lama yang sempat diupload di sesi sebelumnya — karena has_upload sekali
# True akan TERUS True sepanjang sesi (session_state-nya tidak pernah
# dibersihkan), jadi crawl baru tidak pernah "menang" walau jelas lebih
# baru. Sekarang dipilih berdasarkan aksi mana yang TERAKHIR terjadi
# (upload file, ATAU crawl baru selesai) via
# session_state["_dash_last_source"] — ditulis di titik upload (di atas,
# dan di landing) dan di titik crawl selesai (app.py, tepat setelah
# "hasil"/"klaster" diisi). Kalau cuma salah satu sumber yang ada, sumber
# itu otomatis dipakai; kalau _dash_last_source belum pernah ditandai
# (sesi lama dari sebelum fix ini) fallback ke upload, sama seperti
# perilaku lama, supaya tidak ada perubahan tampilan mendadak untuk sesi
# yang sudah berjalan.

if not has_upload and not has_session:
    # Landing state — belum ada data dari mana pun
    st.markdown("""
    <div class="ais-topbar">
      <div>
        <div class="ais-logo">Dashboard Analisis Isu Strategis</div>
        <div class="ais-subtitle">Analisis Isu Strategis Pengawasan — Pusat Strategi Kebijakan Pengawasan BPKP</div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1,2,1])
    with col2:
        with st.container(key="landing_upload_card"):
            st.markdown("""
            <div style='text-align:center;padding:8px 4px 4px'>
              <div style='font-size:40px;margin-bottom:12px'>📊</div>
              <div style='font-size:16px;font-weight:600;color:inherit;margin-bottom:8px'>Belum Ada Data</div>
              <div style='font-size:12px;color:inherit;opacity:0.6;line-height:1.6;margin-bottom:14px'>
                Jalankan crawl di halaman <b>🔍 Crawl & Analisis</b> terlebih dahulu,<br>
                atau upload file <code>.xlsx</code> hasil crawl sebelumnya di bawah ini.
              </div>
            </div>
            """, unsafe_allow_html=True)
            uploaded_landing = st.file_uploader(
                "Upload file Excel (.xlsx)", type=["xlsx"],
                label_visibility="collapsed", key="uploader_xlsx"
            )
        if uploaded_landing is not None:
            df_raw_baru, meta_baru = load_from_excel(uploaded_landing)
            st.session_state["_dash_upload_df_raw"] = df_raw_baru
            st.session_state["_dash_upload_meta"] = meta_baru
            st.session_state["_dash_upload_file_id"] = getattr(uploaded_landing, "file_id", None) or f"{uploaded_landing.name}_{uploaded_landing.size}"
            st.session_state["_dash_last_source"] = "upload"
            st.rerun()
    st.stop()

# ── LOAD & PROCESS DATA ──────────────────────────────────────
last_source = st.session_state.get("_dash_last_source")
if has_upload and has_session:
    pakai_upload = last_source != "session"
else:
    pakai_upload = has_upload

if pakai_upload:
    # Upload manual dipilih — baik karena cuma ini satu-satunya sumber
    # yang ada, atau karena ini yang paling terakhir dilakukan user
    # dibanding sesi crawl yang sedang aktif. Dibaca dari cache
    # session_state (bukan langsung dari widget) — lihat catatan di atas.
    df_raw = st.session_state["_dash_upload_df_raw"]
    meta = st.session_state["_dash_upload_meta"]
    sumber_data = "upload"
    klaster_meta = []  # narasi klaster lengkap tidak tersedia dari Excel, hanya nama per baris

    # Hidrasi status telaah dari file yang di-upload ke session_state,
    # di-guard per file_id supaya hanya jalan sekali (tidak menimpa telaah
    # baru yang sedang berjalan).
    file_id = st.session_state["_dash_upload_file_id"]
    if st.session_state.get("_review_hydrated_from") != file_id and "StatusReview" in df_raw.columns:
        st.session_state.setdefault("review_klaster", {})
        sudah_direview_df = df_raw[df_raw["StatusReview"] == "Sudah Direview"]
        for nama_klaster, grup in sudah_direview_df.groupby("Klaster"):
            baris = grup.iloc[0]
            st.session_state["review_klaster"].setdefault(nama_klaster, {
                "sektor": baris.get("Sektor", "-"),
                "tema": baris.get("Tema", "-"),
                "topik": baris.get("Topik", "-"),
                "dampak_implikasi_final": baris.get("DampakImplikasiFinal", "-"),
                "gap_pengawasan": baris.get("GapPengawasan", "-"),
                "usulan_pengawasan": baris.get("UsulanPengawasan", "-"),
                "status_review": "Sudah Direview",
            })
        st.session_state["_review_hydrated_from"] = file_id

else:
    # Ambil dari hasil crawl halaman 1
    hasil_list  = st.session_state["hasil"]
    label_sesi  = st.session_state.get("label_isu", "Hasil Crawl")
    klaster_meta = st.session_state.get("klaster", [])
    df_raw = pd.DataFrame([{
        'No':          i + 1,
        'Klaster':     h.get('klaster', '-'),
        'Tanggal':     h.get('tanggal', '-'),
        'Sumber':      h.get('sumber', '-'),
        'Link':        h.get('link', '-'),
        'Judul':       h.get('judul', '-'),
        'Ringkasan':   h.get('ringkasan_isu', '-'),
        'IsuSubisu':   h.get('isu_subisu', '-'),
        'AktorLokasi': h.get('aktor_lokasi', '-'),
        'Tone':        h.get('tone', 'Netral'),
        'Risiko':      h.get('risiko', '-'),
        'TindakLanjut': h.get('area_perhatian', '-'),
        'KondisiPemicu': h.get('kondisi_pemicu', '-'),
        'RelevansiPengawasan': h.get('relevansi_pengawasan', '-'),
        'DimensiPengawasan': ", ".join(h.get('dimensi_pengawasan') or []),
    } for i, h in enumerate(hasil_list)])
    meta = {
        "isu":      label_sesi,
        "generate": "dari sesi crawl aktif",
        "total":    str(len(df_raw)),
        "unit":     "Pusat Strategi Kebijakan Pengawasan BPKP"
    }
    sumber_data = "session"

df, stats = compute_stats(df_raw)

# ── Update Excel (Langkah Kerja 3) ───────────────────────────
# Perlu df_raw & meta yang sudah ter-resolve, jadi diletakkan setelah blok
# upload/session_state di atas (bukan di sidebar awal).
with st.sidebar:
    review_klaster = st.session_state.get("review_klaster", {})
    jml_direview = len(review_klaster)

    st.markdown(f"""
    <div style='
        border:1px solid rgba(245,166,35,0.4);
        border-left:4px solid #F5A623;
        border-radius:8px;
        background:rgba(245,166,35,0.08);
        padding:12px 14px;
        margin:14px 0 10px 0;
    '>
      <div style='font-size:13px;font-weight:700;color:#F5A623;margin-bottom:4px'>📥 Update Excel</div>
      <div style='font-size:11px;opacity:0.75;line-height:1.4'>
        {f"{jml_direview} klaster sudah ditelaah dan siap ditulis ke Excel." if jml_direview > 0 else "Belum ada klaster yang ditelaah. Isi form telaah di Tab Klasterisasi Isu."}
      </div>
    </div>
    """, unsafe_allow_html=True)

    if st.button("📊 Generate Excel Terbaru", use_container_width=True, type="primary"):
        # Dicek dari sumber_data (hasil resolusi df_raw/meta di atas),
        # bukan dari nilai widget "uploaded" — widget itu bisa kosong
        # walau datanya berasal dari upload (lihat catatan panjang soal
        # has_upload di bagian SIDEBAR di atas).
        if sumber_data == "session":
            # Sumber: sesi crawl aktif — hasil_list mentah sudah dalam
            # format dict yang dipahami buat_excel.
            sumber_baris = st.session_state["hasil"]
            label_file = st.session_state.get("label_isu", "Hasil Crawl")
        else:
            # Sumber: upload Excel manual — konversi dataframe (df_raw,
            # hasil load_from_excel) balik ke format dict per baris.
            sumber_baris = [
                {
                    "klaster": r.get("Klaster", "-"),
                    "tanggal": r.get("Tanggal", "-"),
                    "sumber": r.get("Sumber", "-"),
                    "link": r.get("Link", "-"),
                    "judul": r.get("Judul", "-"),
                    "ringkasan_isu": r.get("Ringkasan", "-"),
                    "isu_subisu": r.get("IsuSubisu", "-"),
                    "aktor_lokasi": r.get("AktorLokasi", "-"),
                    "tone": r.get("Tone", "Netral"),
                    "risiko": r.get("Risiko", "-"),
                    "area_perhatian": r.get("TindakLanjut", "-"),
                    "kondisi_pemicu": r.get("KondisiPemicu", "-"),
                    "relevansi_pengawasan": r.get("RelevansiPengawasan", "-"),
                }
                for r in df_raw.to_dict("records")
            ]
            label_file = meta.get("isu", "Hasil Upload")

        hasil_terbaru = []
        for h in sumber_baris:
            h2 = dict(h)
            nama_klaster = h2.get("klaster", "-")
            review = review_klaster.get(nama_klaster)
            if review:
                h2["sektor"] = review.get("sektor", "-")
                h2["tema"] = review.get("tema", "-")
                h2["topik"] = review.get("topik", "-")
                h2["dampak_implikasi_final"] = review.get("dampak_implikasi_final", "-")
                h2["gap_pengawasan"] = review.get("gap_pengawasan", "-")
                h2["usulan_pengawasan"] = review.get("usulan_pengawasan", "-")
                h2["status_review"] = review.get("status_review", "Belum Direview")
            hasil_terbaru.append(h2)

        # buat_excel() dari app.py (scope sama, dieksekusi via exec())
        excel_bytes = buat_excel(hasil_terbaru, label_file)
        st.download_button(
            "⬇️ Download Excel",
            data=excel_bytes,
            file_name=f"MediaCrawl_AIS_{str(label_file).replace(' ','_')}_telaah.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    st.markdown("<div style='font-size:10px;opacity:0.55;margin:10px 0 4px 2px'>Unggah hasil telaah Anda di sini ⬇️</div>", unsafe_allow_html=True)
    st.link_button(
        "📁 Drive Hasil Telaah",
        "https://drive.google.com/drive/u/0/folders/1hRyMkpe6TVgaSDs8uXbHZfRkrmZPcxDE",
        use_container_width=True,
    )

# Info banner sumber data
if sumber_data == "session":
    st.info(f"📡 Menampilkan hasil crawl sesi ini: **{meta.get('isu','—')}** · {meta.get('total','—')} artikel — Upload file Excel di sidebar untuk mengganti data.")

# Apply filters
df_filtered = df[
    (df['Tone'].isin(filter_tone)) & 
    (df['level_risiko'].isin(filter_risiko))
].reset_index(drop=True)

# ── TOPBAR ───────────────────────────────────────────────────
st.markdown(f"""
<div class="ais-topbar">
  <div>
    <div class="ais-logo">Dashboard Analisis Isu Strategis</div>
    <div class="ais-subtitle">Analisis Isu Strategis Pengawasan — {meta.get('unit','Pusat Strategi Kebijakan Pengawasan BPKP')}</div>
  </div>
  <div>
    <span class="ais-badge">ISU: {meta.get('isu','—').upper()}</span>
    &nbsp;
    <span style='font-size:11px;color:rgba(255,255,255,0.5)'>{meta.get('generate','—')}</span>
  </div>
</div>
""", unsafe_allow_html=True)

# ── TABS ─────────────────────────────────────────────────────
tab1, tab2, tab3, tab4 = st.tabs(["📋 Ikhtisar", "🗂️ Klasterisasi Isu", "📈 Sentimen & Tren", "🔑 Kata Kunci"])


# ════════════════════════════════════════════
# TAB 1 — IKHTISAR
# ════════════════════════════════════════════
with tab1:

    # Stat row — primer (Total & Dominasi Negatif, dua metrik paling
    # actionable untuk pengawasan) di atas, rincian tone sebagai sekunder
    # di bawahnya supaya bobot visual tidak rata.
    c_p1, c_p2 = st.columns(2)
    with c_p1:
        st.markdown(f"""<div class="stat-card-primary" style="border-top:4px solid #F5A623">
          <div class="stat-num-primary">{stats['total']}</div>
          <div class="stat-label-primary">Total Artikel</div>
        </div>""", unsafe_allow_html=True)
    with c_p2:
        st.markdown(f"""<div class="stat-card-primary" style="border-top:4px solid #E74C3C">
          <div class="stat-num-primary" style="color:#E74C3C">{stats['pct_neg']}%</div>
          <div class="stat-label-primary">Dominasi Negatif</div>
        </div>""", unsafe_allow_html=True)

    st.markdown("<div style='margin-top:8px'></div>", unsafe_allow_html=True)

    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown(f"""<div class="stat-card-secondary">
          <div class="stat-num-secondary" style="color:#E74C3C">{stats['negatif']}</div>
          <div class="stat-label-secondary">🔴 Negatif</div>
        </div>""", unsafe_allow_html=True)
    with c2:
        st.markdown(f"""<div class="stat-card-secondary">
          <div class="stat-num-secondary" style="color:#7F8C8D">{stats['netral']}</div>
          <div class="stat-label-secondary">⚪ Netral</div>
        </div>""", unsafe_allow_html=True)
    with c3:
        st.markdown(f"""<div class="stat-card-secondary">
          <div class="stat-num-secondary" style="color:#27AE60">{stats['positif']}</div>
          <div class="stat-label-secondary">🟢 Positif</div>
        </div>""", unsafe_allow_html=True)

    st.markdown("<div style='margin-top:16px'></div>", unsafe_allow_html=True)

    # Spotlight — artikel paling relevan dengan topik crawl.
    # Prioritas: (1) Negatif + judul relevan, (2) Negatif apapun, (3) artikel pertama
    def skor_relevansi(row, topik: str) -> int:
        """Hitung relevansi judul terhadap topik crawl."""
        kata_topik = set(topik.lower().replace("-","").replace("_"," ").split())
        kata_judul = set(str(row['Judul']).lower().split())
        # Jumlah kata topik yang muncul di judul
        match = len(kata_topik & kata_judul)
        # Bonus jika Negatif
        tone_bonus = 3 if str(row['Tone']) == 'Negatif' else 0
        # Bonus panjang ringkasan (konten lebih kaya)
        content_bonus = min(2, len(str(row['Ringkasan'])) // 100)
        return match * 2 + tone_bonus + content_bonus

    topik_crawl = meta.get('isu', '')
    df_score = df.copy()
    df_score['relevansi'] = df_score.apply(lambda r: skor_relevansi(r, topik_crawl), axis=1)
    df_score = df_score.sort_values('relevansi', ascending=False)

    spotlight_idx = df_score.index[0] if len(df_score) > 0 else None

    if len(df_score) > 0:
        spotlight = df_score.iloc[0]
        tone_spotlight = str(spotlight['Tone'])
        tone_color = {'Negatif': '#E74C3C', 'Netral': '#95A5A6', 'Positif': '#27AE60'}.get(tone_spotlight, '#95A5A6')
        judul_spotlight, sumber_spotlight = pisahkan_sumber_judul(spotlight['Judul'])
        st.markdown(f"""
        <div class="spotlight-box">
          <div class="spotlight-eyebrow">⚡ Isu Prioritas — Paling Relevan & Signifikan</div>
          <div class="spotlight-title">{pill_sumber_html(sumber_spotlight)}{judul_spotlight}</div>
          <div class="spotlight-body">{spotlight['Ringkasan']}<br><br>
            <strong style="color:rgba(255,255,255,0.9)">Risiko:</strong> {spotlight['Risiko']}
          </div>
        </div>
        """, unsafe_allow_html=True)

    # Artikel risiko tinggi — ditaruh sebelum chart supaya insight yang
    # paling actionable terlihat duluan, baru statistik pendukungnya.
    # Artikel yang sudah muncul di Spotlight di-exclude supaya tidak dobel.
    st.markdown("**Artikel Risiko Tinggi — Perlu Perhatian**")
    df_tinggi = df[df['level_risiko'] == 'Tinggi']
    if spotlight_idx is not None:
        df_tinggi = df_tinggi[df_tinggi.index != spotlight_idx]
    df_tinggi = df_tinggi.head(3)
    if len(df_tinggi) == 0:
        df_tinggi = df[df['Tone'] == 'Negatif']
        if spotlight_idx is not None:
            df_tinggi = df_tinggi[df_tinggi.index != spotlight_idx]
        df_tinggi = df_tinggi.head(3)

    if len(df_tinggi) == 0:
        st.info("Tidak ada artikel risiko tinggi/negatif lain di luar Isu Prioritas pada periode ini.")
    else:
        cols_tinggi = st.columns(min(3, len(df_tinggi)))
        for i, (_, row) in enumerate(df_tinggi.iterrows()):
            with cols_tinggi[i]:
                tone_class = str(row['Tone']).lower()
                aktor_short = str(row['AktorLokasi'])[:40]
                judul_bersih, sumber_row = pisahkan_sumber_judul(row['Judul'])
                judul_disp = judul_bersih[:80] + ('…' if len(judul_bersih) > 80 else '')
                st.markdown(f"""
                <div class="issue-card issue-card-highlight {tone_class}">
                  <div class="issue-title">{pill_sumber_html(sumber_row)}{judul_disp}</div>
                  <div class="issue-sub">{str(row['IsuSubisu'])}</div>
                  <div class="issue-summary">{str(row['Ringkasan'])[:160]}…</div>
                  <div style='margin-top:6px'>
                    <span class="badge badge-{tone_class}">{row['Tone']}</span>
                    <span class="badge badge-aktor">👤 {aktor_short}</span>
                  </div>
                </div>
                """, unsafe_allow_html=True)

    st.markdown("<div style='margin-top:8px'></div>", unsafe_allow_html=True)
    st.markdown("---")

    # Tone bar + frekuensi subisu
    col_left, col_right = st.columns(2)

    with col_left:
        st.markdown("**Distribusi Sentimen Pemberitaan**")
        
        # Visual tone bar
        pn = stats['pct_neg']; pt = stats['pct_net']; pp = stats['pct_pos']
        st.markdown(f"""
        <div style='height:12px;border-radius:6px;overflow:hidden;display:flex;margin-bottom:8px'>
          <div style='width:{pn}%;background:#E74C3C'></div>
          <div style='width:{pt}%;background:#BDC3C7'></div>
          <div style='width:{pp}%;background:#27AE60'></div>
        </div>
        <div style='display:flex;gap:16px;font-size:11px;color:inherit;opacity:0.7'>
          <span>🔴 Negatif {pn}%</span>
          <span>⚪ Netral {pt}%</span>
          <span>🟢 Positif {pp}%</span>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("<div style='margin-top:16px'></div>", unsafe_allow_html=True)

        # Level risiko breakdown
        st.markdown("**Level Risiko Artikel**")
        tinggi = stats['tinggi']; sedang = stats['sedang']
        rendah = stats['total'] - tinggi - sedang
        for label, count, color in [("Tinggi", tinggi, "#E74C3C"), ("Sedang", sedang, "#F5A623"), ("Rendah", rendah, "#27AE60")]:
            pct = round(count/stats['total']*100) if stats['total'] else 0
            st.markdown(f"""
            <div style='display:flex;align-items:center;gap:8px;margin-bottom:6px'>
              <div style='font-size:11px;color:inherit;opacity:0.7;width:50px'>{label}</div>
              <div style='flex:1;height:14px;background:rgba(128,128,128,0.15);border-radius:3px;overflow:hidden'>
                <div style='width:{pct}%;height:100%;background:{color};border-radius:3px'></div>
              </div>
              <div style='font-size:10px;font-family:monospace;color:#95A5A6;width:24px'>{count}</div>
            </div>
            """, unsafe_allow_html=True)

    with col_right:
        st.markdown("**Sebaran Risiko per Aktor/Lokasi**")
        st.caption("Diurutkan berdasarkan jumlah artikel risiko Tinggi — menunjukkan instansi/aktor yang paling sering terkait isu berisiko")
        aktor_risiko = risiko_per_aktor(df, top_n=8)

        if not aktor_risiko:
            st.info("Belum ada data Aktor/Lokasi yang bisa dianalisis.")
        else:
            max_total = max(a['total'] for a in aktor_risiko)
            for a in aktor_risiko:
                label = (a['aktor'][:32]+'…') if len(a['aktor'])>32 else a['aktor']
                pct_tinggi = round(a['tinggi'] / max_total * 100)
                pct_sedang = round(a['sedang'] / max_total * 100)
                pct_rendah = round(a['rendah'] / max_total * 100)
                st.markdown(f"""
                <div style='margin-bottom:10px'>
                  <div style='display:flex;justify-content:space-between;margin-bottom:3px'>
                    <span style='font-size:11px;color:inherit;opacity:0.85;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:220px' title='{a["aktor"]}'>{label}</span>
                    <span style='font-size:10px;font-family:monospace;color:inherit;opacity:0.5'>{a['total']} artikel</span>
                  </div>
                  <div style='display:flex;height:12px;background:rgba(128,128,128,0.15);border-radius:3px;overflow:hidden'>
                    <div style='width:{pct_tinggi}%;height:100%;background:#E74C3C' title='Tinggi: {a["tinggi"]}'></div>
                    <div style='width:{pct_sedang}%;height:100%;background:#F5A623' title='Sedang: {a["sedang"]}'></div>
                    <div style='width:{pct_rendah}%;height:100%;background:#27AE60' title='Rendah: {a["rendah"]}'></div>
                  </div>
                </div>
                """, unsafe_allow_html=True)
            st.markdown("""
            <div style='display:flex;gap:14px;margin-top:8px;font-size:10px;color:inherit;opacity:0.7'>
              <span>🔴 Tinggi</span><span>🟡 Sedang</span><span>🟢 Rendah</span>
            </div>
            """, unsafe_allow_html=True)


# ════════════════════════════════════════════
# TAB 2 — DAFTAR ISU
# ════════════════════════════════════════════
with tab2:
    st.markdown(f"**{len(df_filtered)} artikel** ditampilkan berdasarkan filter aktif")

    if len(df_filtered) == 0:
        st.info("Tidak ada artikel yang sesuai filter. Ubah filter di sidebar.")
    else:
        # Split: list kiri, detail kanan
        col_list, col_detail = st.columns([5, 4])

        with col_list:
            # Pilih artikel — default None (belum ada yang diklik), BUKAN 0,
            # supaya tidak ada artikel yang ke-highlight "terpilih" atau
            # panel kanan menampilkan detail artikel acak sebelum user
            # benar-benar mengklik "Lihat detail →".
            selected_idx = st.session_state.get('selected_idx')

            ada_klaster = 'Klaster' in df_filtered.columns and (df_filtered['Klaster'] != '-').any()

            def render_artikel_item(i, row, compact=False):
                tone_class = str(row['Tone']).lower()
                is_selected = (i == selected_idx)
                border_style = "border:2px solid rgba(99,179,237,0.8);" if is_selected else "border:1px solid rgba(128,128,128,0.2);"
                bg_style = "background:rgba(99,179,237,0.08);" if is_selected else "background:transparent;"
                varian_class = "issue-card-member" if compact else ""

                judul_bersih, sumber_row = pisahkan_sumber_judul(row['Judul'])
                judul_short = judul_bersih[:75]+'…' if len(judul_bersih)>75 else judul_bersih

                btn_key = f"artikel_{i}"
                # Card HTML + tombol "Lihat detail" dibungkus SATU
                # st.container(key=...) supaya keduanya benar-benar
                # bertetangga dalam satu elemen DOM (bukan dua widget
                # terpisah yang cuma kebetulan berdekatan) — sebelumnya ini
                # bikin gak jelas tombol itu milik card di atas atau malah
                # nempel ke card berikutnya. CSS di bawah menyambungkan
                # visualnya jadi satu unit: card persegi di atas, tombol
                # jadi strip footer nempel tanpa celah di bawahnya.
                with st.container(key=f"artikel_card_{i}"):
                    st.markdown(f"""
                    <div class="issue-card artikel-card-attached {varian_class} {tone_class}" style="{border_style}{bg_style}">
                      <div style='display:flex;justify-content:space-between;align-items:flex-start;gap:8px'>
                        <div>
                          <div class="issue-title">{pill_sumber_html(sumber_row, compact=compact)}{judul_short}</div>
                          <div class="issue-sub">{str(row['IsuSubisu'])}</div>
                        </div>
                        <span class="badge badge-{tone_class}" style='flex-shrink:0'>{row['Tone']}</span>
                      </div>
                      <div style='margin-top:6px;font-size:10px;color:inherit;opacity:0.5;font-family:monospace'>
                        📅 {row['Tanggal']} &nbsp;·&nbsp; 🏢 {str(row['AktorLokasi'])[:40]}
                      </div>
                    </div>
                    """, unsafe_allow_html=True)

                    if st.button(f"Lihat detail →", key=btn_key, use_container_width=True):
                        st.session_state['selected_idx'] = i
                        st.rerun()

            narasi_klaster = {}  # selalu didefinisikan, diisi jika ada_klaster True

            if not ada_klaster:
                # Fallback: tampilan datar (Excel lama tanpa kolom Klaster)
                for i, (_, row) in enumerate(df_filtered.iterrows()):
                    render_artikel_item(i, row)
            else:
                # Bangun lookup narasi klaster (jika tersedia dari sesi crawl aktif)
                narasi_klaster = {k.get('nama', '-'): k for k in klaster_meta} if klaster_meta else {}

                df_filtered_idx = df_filtered.reset_index(drop=True)
                klaster_order = [k.get('nama','-') for k in klaster_meta] if klaster_meta else None
                nama_unik = df_filtered_idx['Klaster'].fillna('-').unique().tolist()
                if klaster_order:
                    nama_terurut = [n for n in klaster_order if n in nama_unik] + [n for n in nama_unik if n not in klaster_order]
                else:
                    nama_terurut = sorted(nama_unik, key=lambda n: -(df_filtered_idx['Klaster']==n).sum())

                for nama in nama_terurut:
                    sub_idx = df_filtered_idx[df_filtered_idx['Klaster'] == nama]
                    jumlah = len(sub_idx)
                    if jumlah == 0:
                        continue

                    tone_dom = sub_idx['Tone'].value_counts().idxmax() if jumlah else 'Netral'
                    dom_color = {'Negatif':'#E74C3C','Netral':'#95A5A6','Positif':'#27AE60'}.get(tone_dom,'#95A5A6')

                    label_expander = f"🗂️ **{nama}**  ·  :gray[{jumlah} artikel]"
                    # Semua klaster tertutup by default (bukan cuma klaster
                    # pertama) — user baru tidak langsung dihadapkan detail
                    # klaster + panel artikel di kanan sebelum sempat
                    # memindai daftar klaster yang ada.
                    with st.expander(label_expander, expanded=False):
                        info_klaster = narasi_klaster.get(nama)

                        # Fallback: rekonstruksi field klaster dari data
                        # artikel jika narasi klaster tidak tersedia (mis.
                        # sumber data dari upload Excel).
                        if not info_klaster:
                            def _ambil_unik(kolom):
                                if kolom not in sub_idx.columns:
                                    return "-"
                                nilai = sub_idx[kolom].dropna()
                                nilai = nilai[nilai != "-"]
                                return nilai.iloc[0] if len(nilai) else "-"

                            # Dimensi Pengawasan disimpan di Excel sebagai
                            # string dipisah koma (mis. "Anti-Korupsi, Control")
                            # — pecah balik jadi list, dan saring ulang ke
                            # daftar resmi supaya kalau file diedit manual
                            # (typo, format lain) tidak ikut nyasar ke UI,
                            # cukup diam-diam diabaikan.
                            DIMENSI_PENGAWASAN_VALID = {"Governance", "Risk", "Control", "Compliance", "Anti-Korupsi", "Debottlenecking"}
                            dimensi_mentah = _ambil_unik('DimensiPengawasan')
                            dimensi_upload = (
                                [d.strip() for d in dimensi_mentah.split(",") if d.strip() in DIMENSI_PENGAWASAN_VALID]
                                if dimensi_mentah and dimensi_mentah != "-" else []
                            )

                            info_klaster = {
                                "kondisi_pemicu":       _ambil_unik('KondisiPemicu'),
                                "risiko":               _ambil_unik('Risiko'),
                                "area_perhatian":       _ambil_unik('TindakLanjut'),
                                "relevansi_pengawasan": _ambil_unik('RelevansiPengawasan'),
                                "dimensi_pengawasan":   dimensi_upload,
                            }

                        # Tag dimensi pengawasan (GRCC AnCoDe) — cuma render
                        # baris ini kalau memang ada tag yang lolos sanitasi di
                        # klasterisasi_isu_deepseek(); klaster tanpa dimensi
                        # jelas (list kosong) tidak dipaksa tampil baris kosong.
                        dimensi_list = info_klaster.get('dimensi_pengawasan') or []
                        dimensi_html = "".join(f'<span class="badge badge-dimensi">{d}</span>' for d in dimensi_list)
                        dimensi_row = (
                            f"<div style='font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:#F5A623;opacity:0.85;margin-top:10px;margin-bottom:5px'>Dimensi Pengawasan</div>"
                            f"<div>{dimensi_html}</div>"
                        ) if dimensi_list else ""

                        st.markdown(f"""
                        <div style='
                            border-left:6px solid #F5A623;
                            border-radius:0 10px 10px 0;
                            background:linear-gradient(135deg, rgba(245,166,35,0.10), rgba(245,166,35,0.03));
                            padding:16px 20px;
                            margin-bottom:4px;
                            box-shadow:0 2px 8px rgba(0,0,0,0.15);
                        '>
                          <div style='display:flex;align-items:center;gap:6px;margin-bottom:12px'>
                            <span style='font-size:10px;font-weight:800;letter-spacing:.1em;color:#F5A623;text-transform:uppercase;font-family:monospace'>📁 INDUK KLASTER</span>
                            <span style='font-size:10px;opacity:0.4'>·</span>
                            <span style='font-size:10px;opacity:0.55'>{jumlah} artikel anggota</span>
                          </div>
                          <div style='font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:#F5A623;opacity:0.85;margin-bottom:3px'>Kondisi / Pemicu</div>
                          <div style='font-size:14px;line-height:1.6;margin-bottom:10px;font-weight:600'>{info_klaster.get('kondisi_pemicu','-')}</div>
                          <div style='font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:#F5A623;opacity:0.85;margin-bottom:3px'>Risiko</div>
                          <div style='font-size:14px;line-height:1.6;margin-bottom:10px;font-weight:600'>{info_klaster.get('risiko','-')}</div>
                          <div style='font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:#F5A623;opacity:0.85;margin-bottom:3px'>Area Perhatian</div>
                          <div style='font-size:14px;line-height:1.6;margin-bottom:10px;font-weight:600'>{info_klaster.get('area_perhatian','-')}</div>
                          <div style='font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:#F5A623;opacity:0.85;margin-bottom:3px'>Relevansi Pengawasan BPKP</div>
                          <div style='font-size:14px;line-height:1.6;font-weight:600'>{info_klaster.get('relevansi_pengawasan','-')}</div>
                          {dimensi_row}
                        </div>
                        """, unsafe_allow_html=True)

                        # Tombol buka telaah — form lengkap ada di panel kanan
                        # (col_detail), diaktifkan lewat panel_mode="telaah".
                        review_tersimpan = st.session_state.get("review_klaster", {}).get(nama, {})
                        sudah_direview = bool(review_tersimpan.get("status_review") == "Sudah Direview")
                        badge_review = "🟢 Sudah Direview" if sudah_direview else "⚪ Belum Direview"

                        c_badge, c_btn = st.columns([2, 1.6])
                        with c_badge:
                            st.markdown(f"<div style='font-size:11px;font-weight:700;opacity:0.7;padding-top:8px'>{badge_review}</div>", unsafe_allow_html=True)
                        with c_btn:
                            if st.button("✏️ Telaah klaster ini →", key=f"buka_telaah_{nama}", use_container_width=True):
                                st.session_state["panel_mode"] = "telaah"
                                st.session_state["selected_klaster"] = nama
                                st.rerun()

                        st.markdown("<div style='display:flex;align-items:center;gap:10px;margin:14px 0 10px 4px'><span style='font-size:10px;font-weight:700;letter-spacing:.08em;opacity:0.45;text-transform:uppercase;white-space:nowrap'>↳ Artikel Anggota</span><div style='flex:1;height:1px;background:rgba(128,128,128,0.25)'></div></div>", unsafe_allow_html=True)

                        st.markdown("<div style='margin-left:14px;border-left:1px dashed rgba(128,128,128,0.25);padding-left:14px'>", unsafe_allow_html=True)
                        for i, row in sub_idx.iterrows():
                            render_artikel_item(i, row, compact=True)
                        st.markdown("</div>", unsafe_allow_html=True)

        with col_detail:
            panel_mode = st.session_state.get("panel_mode", "detail")

            with st.container(key="panel_kanan"):
                if panel_mode == "telaah" and st.session_state.get("selected_klaster"):
                    nama_aktif = st.session_state["selected_klaster"]
                    review_key = f"review_{nama_aktif}"
                    review_tersimpan = st.session_state.get("review_klaster", {}).get(nama_aktif, {})

                    # Ambil draft risiko AI untuk klaster ini sebagai starting
                    # point Dampak/Implikasi, dari narasi_klaster jika tersedia
                    risiko_draft = narasi_klaster.get(nama_aktif, {}).get("risiko", "")
                    if not risiko_draft:
                        sub_match = df_filtered[df_filtered['Klaster'] == nama_aktif]
                        risiko_vals = sub_match['Risiko'].dropna() if 'Risiko' in sub_match.columns else pd.Series([])
                        risiko_draft = risiko_vals.iloc[0] if len(risiko_vals) else ""

                    st.markdown(f"""
                    <div style='display:flex;align-items:center;gap:6px;margin-bottom:10px;
                        font-family:"JetBrains Mono",monospace;font-size:10px;letter-spacing:0.08em;
                        color:#F5A623;font-weight:700;text-transform:uppercase;'>
                      ✏️ TELAAH KLASTER
                    </div>
                    <div style='font-size:14px;font-weight:700;margin-bottom:12px;line-height:1.4'>{nama_aktif}</div>
                    """, unsafe_allow_html=True)

                    if st.button("← Tutup telaah, lihat detail artikel", key="tutup_telaah", use_container_width=True):
                        st.session_state["panel_mode"] = "detail"
                        st.rerun()

                    st.markdown("<hr style='border:none;border-top:1px solid rgba(245,166,35,0.2);margin:10px 0'>", unsafe_allow_html=True)

                    sektor_list = list(STRUKTUR_APP.keys())
                    sektor_default = review_tersimpan.get("sektor", sektor_list[0])
                    sektor_idx = sektor_list.index(sektor_default) if sektor_default in sektor_list else 0

                    col_sektor, col_tema, col_topik = st.columns(3)
                    with col_sektor:
                        sektor_pilih = st.selectbox("Sektor", sektor_list, index=sektor_idx, key=f"{review_key}_sektor")

                    tema_list = list(STRUKTUR_APP.get(sektor_pilih, {}).keys())
                    tema_default = review_tersimpan.get("tema", tema_list[0] if tema_list else None)
                    tema_idx = tema_list.index(tema_default) if tema_default in tema_list else 0
                    with col_tema:
                        tema_pilih = st.selectbox("Tema", tema_list, index=tema_idx, key=f"{review_key}_tema") if tema_list else None

                    topik_list = STRUKTUR_APP.get(sektor_pilih, {}).get(tema_pilih, []) if tema_pilih else []
                    topik_default = review_tersimpan.get("topik", topik_list[0] if topik_list else None)
                    topik_idx = topik_list.index(topik_default) if topik_default in topik_list else 0
                    with col_topik:
                        topik_pilih = st.selectbox("Topik", topik_list, index=topik_idx, key=f"{review_key}_topik") if topik_list else None

                    dampak_default = review_tersimpan.get("dampak_implikasi_final") or risiko_draft
                    dampak_pilih = st.text_area("Dampak / Implikasi (sempurnakan draf AI)", value=dampak_default, key=f"{review_key}_dampak", height=110)

                    gap_pilih = st.text_area("Gap Pengawasan", value=review_tersimpan.get("gap_pengawasan", ""), key=f"{review_key}_gap", height=90,
                                               placeholder="Apa yang belum tercakup dalam pengawasan eksisting BPKP terhadap isu ini?")

                    usulan_pilih = st.text_area("Usulan Pengawasan", value=review_tersimpan.get("usulan_pengawasan", ""), key=f"{review_key}_usulan", height=90,
                                                  placeholder="Usulan lingkup/metodologi pengawasan untuk mengakomodir isu ini")

                    if st.button("💾 Submit Telaah", key=f"{review_key}_submit", use_container_width=True, type="primary"):
                        if "review_klaster" not in st.session_state:
                            st.session_state["review_klaster"] = {}
                        st.session_state["review_klaster"][nama_aktif] = {
                            "sektor": sektor_pilih,
                            "tema": tema_pilih or "-",
                            "topik": topik_pilih or "-",
                            "dampak_implikasi_final": dampak_pilih,
                            "gap_pengawasan": gap_pilih,
                            "usulan_pengawasan": usulan_pilih,
                            "status_review": "Sudah Direview",
                        }
                        st.success(f"Telaah tersimpan. Klik 'Update Excel' di sidebar untuk menulis ke file.")
                        st.rerun()

                else:
                    idx = st.session_state.get('selected_idx')
                    if idx is not None and idx < len(df_filtered):
                        row = df_filtered.iloc[idx]
                        tone_class = str(row['Tone']).lower()
                        judul_bersih, sumber_row = pisahkan_sumber_judul(row['Judul'])

                        # Aktor/Lokasi dipecah jadi pill per nama (bukan satu
                        # blob teks) supaya tiap pihak yang terlibat gampang
                        # dipindai satu-satu, dan dikasih ikon 👤 + warna
                        # slate yang beda dari badge topik (indigo) di
                        # atasnya — dua-duanya kelihatan sama sebelum ini
                        # padahal beda level informasi (kategori isu vs.
                        # pihak yang terlibat).
                        daftar_aktor = [a.strip() for a in str(row['AktorLokasi']).split(',') if a.strip() and a.strip() != '-']
                        aktor_pills = "".join(f'<span class="badge badge-aktor">👤 {a}</span>' for a in daftar_aktor) or '<span class="badge badge-aktor">👤 -</span>'

                        st.markdown(f"""
                        <div style='
                            display:flex;align-items:center;gap:6px;margin-bottom:14px;
                            font-family:"JetBrains Mono",monospace;font-size:10px;
                            letter-spacing:0.08em;color:#F5A623;font-weight:700;
                            text-transform:uppercase;
                        '>
                          📋 Detail Analisis
                        </div>

                        <div style='display:flex;justify-content:space-between;align-items:flex-start;gap:8px;margin-bottom:10px'>
                          <div style='font-size:14px;font-weight:700;color:inherit;line-height:1.4;flex:1'>{pill_sumber_html(sumber_row)}{judul_bersih}</div>
                          <span class="badge badge-{tone_class}" style='font-size:11px;padding:3px 8px;flex-shrink:0'>{row['Tone']}</span>
                        </div>

                        <div style='margin-bottom:10px'>
                          <span class="badge badge-topik">🏷️ {row['IsuSubisu']}</span>
                        </div>

                        <div class="detail-label" style='margin-bottom:6px'>Aktor / Lokasi Terkait</div>
                        <div style='display:flex;gap:6px;flex-wrap:wrap;margin-bottom:12px'>
                          {aktor_pills}
                        </div>

                        <hr style='border:none;border-top:1px solid rgba(245,166,35,0.2);margin:10px 0'>

                        <div class="detail-section">
                          <div class="detail-label">Ringkasan Isu</div>
                          <div class="detail-text">{row['Ringkasan']}</div>
                        </div>

                        <hr style='border:none;border-top:1px solid rgba(245,166,35,0.2);margin:10px 0'>
                        <div style='font-size:11px;color:inherit;display:flex;align-items:center;gap:8px;flex-wrap:wrap'>
                          <span style='opacity:0.5'>📅 {row['Tanggal']}</span>
                          <span style='opacity:0.3'>·</span>
                          <a href="{row['Link']}" target="_blank" class="link-artikel-asli">🔗 Buka artikel asli →</a>
                        </div>
                        """, unsafe_allow_html=True)
                    else:
                        st.info("Pilih artikel di kiri untuk melihat detail.")


# ════════════════════════════════════════════
# TAB 3 — TONE & TREN
# ════════════════════════════════════════════
with tab3:
    # Catatan analitis ditaruh paling atas — ini kesimpulan yang dicari
    # user, tabel & chart di bawah adalah rincian pendukungnya.
    neg_issues = df[df['Tone']=='Negatif']['IsuSubisu'].value_counts().head(3).index.tolist()
    pos_issues = df[df['Tone']=='Positif']['IsuSubisu'].value_counts().head(2).index.tolist()

    st.markdown("**Catatan Analitis**")
    col_a, col_b, col_c = st.columns(3)
    with col_a:
        st.markdown(f"""
        <div style='background:rgba(231,76,60,0.12);border-radius:6px;padding:12px;border-left:3px solid #E74C3C'>
          <div style='font-size:10px;font-weight:700;color:#E74C3C;margin-bottom:6px;text-transform:uppercase;letter-spacing:.08em'>Isu Dominan Negatif</div>
          <div style='font-size:11px;color:inherit;line-height:1.6'>{"<br>".join(f"• {x}" for x in neg_issues) if neg_issues else "—"}</div>
        </div>
        """, unsafe_allow_html=True)
    with col_b:
        st.markdown(f"""
        <div style='background:rgba(245,166,35,0.12);border-radius:6px;padding:12px;border-left:3px solid #F5A623'>
          <div style='font-size:10px;font-weight:700;color:#c47d0a;margin-bottom:6px;text-transform:uppercase;letter-spacing:.08em'>Volume Pemberitaan</div>
          <div style='font-size:11px;color:inherit;line-height:1.6'>
            Total <strong>{stats['total']}</strong> artikel dalam periode ini.<br>
            {stats['negatif']} negatif ({stats['pct_neg']}%) menunjukkan tekanan pemberitaan yang perlu diwaspadai.
          </div>
        </div>
        """, unsafe_allow_html=True)
    with col_c:
        st.markdown(f"""
        <div style='background:rgba(39,174,96,0.12);border-radius:6px;padding:12px;border-left:3px solid #27AE60'>
          <div style='font-size:10px;font-weight:700;color:#27AE60;margin-bottom:6px;text-transform:uppercase;letter-spacing:.08em'>Isu Bernada Positif</div>
          <div style='font-size:11px;color:inherit;line-height:1.6'>{"<br>".join(f"• {x}" for x in pos_issues) if pos_issues else "—"}</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("---")

    col_tbl, col_chart = st.columns([3, 2])

    with col_tbl:
        st.markdown("**Sentimen per Subisu**")
        
        tone_table = df.groupby(['IsuSubisu','Tone']).size().unstack(fill_value=0)
        for col_name in ['Negatif','Netral','Positif']:
            if col_name not in tone_table.columns:
                tone_table[col_name] = 0
        tone_table = tone_table[['Negatif','Netral','Positif']]
        tone_table['Total'] = tone_table.sum(axis=1)
        tone_table['Tone Dominan'] = tone_table[['Negatif','Netral','Positif']].idxmax(axis=1)
        tone_table = tone_table.sort_values('Total', ascending=False)

        for subisu, row_t in tone_table.iterrows():
            dom = row_t['Tone Dominan']
            dom_color = {'Negatif':'#E74C3C','Netral':'#7F8C8D','Positif':'#27AE60'}.get(dom,'#7F8C8D')
            label = (subisu[:45]+'…') if len(subisu)>45 else subisu
            st.markdown(f"""
            <div style='display:flex;align-items:center;justify-content:space-between;
                        padding:8px 10px;margin-bottom:4px;background:rgba(128,128,128,0.06);
                        border:1px solid rgba(128,128,128,0.15);border-radius:5px;border-left:3px solid {dom_color}'>
              <div style='font-size:11px;font-weight:600;color:inherit;flex:1'>{label}</div>
              <div style='display:flex;gap:10px;font-size:10px;font-family:monospace'>
                <span style='color:#E74C3C'>N:{int(row_t['Negatif'])}</span>
                <span style='color:#95A5A6'>T:{int(row_t['Netral'])}</span>
                <span style='color:#27AE60'>P:{int(row_t['Positif'])}</span>
                <span style='color:inherit;font-weight:700'>{int(row_t['Total'])}</span>
              </div>
            </div>
            """, unsafe_allow_html=True)

    with col_chart:
        st.markdown("**Distribusi Sentimen Keseluruhan**")
        
        import streamlit as st
        tone_data = df['Tone'].value_counts()
        colors_map = {'Negatif':'#E74C3C','Netral':'#BDC3C7','Positif':'#27AE60'}
        
        # Visual bar chart manual — lebih clean dari st.bar_chart default
        for tone_val, count in tone_data.items():
            pct = round(count/stats['total']*100)
            color = colors_map.get(str(tone_val),'#BDC3C7')
            st.markdown(f"""
            <div style='margin-bottom:12px'>
              <div style='display:flex;justify-content:space-between;font-size:11px;margin-bottom:4px'>
                <span style='font-weight:600;color:{color}'>{tone_val}</span>
                <span style='font-family:monospace;color:inherit;opacity:0.5'>{count} artikel ({pct}%)</span>
              </div>
              <div style='height:20px;background:rgba(128,128,128,0.15);border-radius:4px;overflow:hidden'>
                <div style='width:{pct}%;height:100%;background:{color};border-radius:4px'></div>
              </div>
            </div>
            """, unsafe_allow_html=True)

        st.markdown("<div style='margin-top:20px'></div>", unsafe_allow_html=True)
        st.markdown("**Distribusi per Tanggal**")
        
        tanggal_counts = df.groupby(['Tanggal','Tone']).size().unstack(fill_value=0)
        if not tanggal_counts.empty:
            st.bar_chart(tanggal_counts, color=["#E74C3C","#BDC3C7","#27AE60"] if all(c in tanggal_counts.columns for c in ['Negatif','Netral','Positif']) else None, height=200)


# ════════════════════════════════════════════
# TAB 4 — KATA KUNCI
# ════════════════════════════════════════════
with tab4:
    col_cloud, col_bar = st.columns(2)

    keywords = extract_keywords(df)
    aktors = extract_aktors(df)

    with col_cloud:
        st.markdown("**Kata Kunci Dominan**")
        if keywords:
            max_freq = keywords[0][1]
            cloud_html = '<div style="display:flex;flex-wrap:wrap;gap:8px;padding:8px 0">'
            for word, freq in keywords:
                size = 11 + (freq / max_freq) * 12
                opacity = 0.4 + (freq / max_freq) * 0.6
                cloud_html += f'<span style="font-size:{size:.0f}px;background:rgba(100,140,180,0.15);border:1px solid rgba(100,140,180,0.3);border-radius:3px;padding:3px 10px;color:inherit;opacity:{opacity:.2f};font-weight:500">{word}</span>'
            cloud_html += '</div>'
            st.markdown(cloud_html, unsafe_allow_html=True)

    with col_bar:
        st.markdown("**Top Kata Kunci — Frekuensi**")
        if keywords:
            max_freq = keywords[0][1]
            for word, freq in keywords[:10]:
                pct = round(freq / max_freq * 100)
                st.markdown(f"""
                <div style='display:flex;align-items:center;gap:8px;margin-bottom:6px'>
                  <div style='font-size:11px;color:inherit;opacity:0.85;width:120px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis'>{word}</div>
                  <div style='flex:1;height:14px;background:rgba(128,128,128,0.15);border-radius:3px;overflow:hidden'>
                    <div style='width:{pct}%;height:100%;background:#1C3D5A;border-radius:3px'></div>
                  </div>
                  <div style='font-size:10px;font-family:monospace;color:inherit;opacity:0.5;width:18px'>{freq}</div>
                </div>
                """, unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("**Aktor & Lembaga yang Disebut**")
    if aktors:
        max_aktor = aktors[0][1]
        aktor_html = '<div style="display:flex;flex-wrap:wrap;gap:8px;padding:4px 0">'
        for aktor, freq in aktors:
            size = 11 + (freq / max_aktor) * 5
            aktor_html += f'<span style="font-size:{size:.0f}px;background:rgba(99,102,241,0.12);border:1px solid rgba(99,102,241,0.3);border-radius:4px;padding:4px 10px;color:#818CF8;font-weight:500">{aktor} <span style=\'font-family:monospace;font-size:9px;opacity:0.7\'>({freq})</span></span>'
        aktor_html += '</div>'
        st.markdown(aktor_html, unsafe_allow_html=True)

    # Export JSON
    st.markdown("---")
    st.markdown("**Export Data**")
    col_exp1, col_exp2 = st.columns(2)
    with col_exp1:
        records = df.to_dict('records')
        json_str = json.dumps({
            "meta": meta,
            "data": [{k: str(v) for k, v in r.items()} for r in records]
        }, ensure_ascii=False, indent=2)
        st.download_button(
            "⬇️ Download JSON (untuk integrasi)",
            data=json_str,
            file_name=f"ais_data_{meta.get('isu','bpkp').replace(' ','_').lower()}.json",
            mime="application/json",
            use_container_width=True
        )
    with col_exp2:
        csv_str = df[['No','Tanggal','Judul','IsuSubisu','Tone','Risiko','TindakLanjut','AktorLokasi']].to_csv(index=False)
        st.download_button(
            "⬇️ Download CSV (ringkasan)",
            data=csv_str,
            file_name=f"ais_ringkasan_{meta.get('isu','bpkp').replace(' ','_').lower()}.csv",
            mime="text/csv",
            use_container_width=True
        )
