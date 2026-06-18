"""
dashboard_ais.py
Halaman Dashboard AIS — Analisis Isu Strategis Pengawasan
Pusat Strategi Kebijakan Pengawasan BPKP

Cara pakai:
- Upload file Excel hasil crawl (.xlsx) via sidebar
- Dashboard otomatis render semua komponen
- Bisa juga load dari ais_data.json jika tersedia
"""

import streamlit as st
import pandas as pd
import json
import io
from collections import Counter
from datetime import datetime

# ── CUSTOM CSS ───────────────────────────────────────────────
st.markdown("""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;600&display=swap');

  html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

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

  /* Stat cards */
  .stat-card {
    background: rgba(128,128,128,0.06); border: 1px solid rgba(128,128,128,0.2);
    border-radius: 8px; padding: 16px;
    text-align: center;
  }
  .stat-num {
    font-family: 'JetBrains Mono', monospace;
    font-size: 28px; font-weight: 700; color: inherit; line-height: 1;
  }
  .stat-label { font-size: 11px; color: inherit; opacity: 0.6; margin-top: 4px; }

  /* Issue card */
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
  .badge-aktor { background: rgba(99,102,241,0.15); color: #818CF8; }

  /* Detail box */
  .detail-section { margin-bottom: 14px; }
  .detail-label {
    font-size: 10px; font-weight: 700;
    text-transform: uppercase; letter-spacing: 0.08em;
    color: inherit; opacity: 0.55; margin-bottom: 4px;
  }
  .detail-text { font-size: 12px; color: inherit; line-height: 1.6; }
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
</style>
""", unsafe_allow_html=True)


# ── DATA LOADER ──────────────────────────────────────────────
def load_from_excel(uploaded_file):
    """Parse Excel output dari pipeline AIS. Mendukung format lama
    (11 kolom, tanpa Klaster) maupun format baru (12 kolom, dengan
    kolom Klaster Isu di posisi ke-2)."""
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
    ada_klaster = "Klaster Isu" in df.columns or df.shape[1] == 12

    if ada_klaster:
        df.columns = ['No','Klaster','Tanggal','Sumber','Link','Judul','Ringkasan','IsuSubisu','AktorLokasi','Tone','Risiko','TindakLanjut']
    else:
        df.columns = ['No','Tanggal','Sumber','Link','Judul','Ringkasan','IsuSubisu','AktorLokasi','Tone','Risiko','TindakLanjut']
        df['Klaster'] = '-'

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
with st.sidebar:
    st.markdown("### 📂 Upload Data")
    st.markdown("<div style='font-size:11px;color:#888;margin-bottom:8px'>Upload file Excel hasil pipeline crawl AIS (.xlsx)</div>", unsafe_allow_html=True)
    
    uploaded = st.file_uploader(
        "Pilih file Excel", type=["xlsx"],
        label_visibility="collapsed"
    )
    
    st.markdown("---")
    st.markdown("### 🔍 Filter")
    filter_tone = st.multiselect(
        "Filter Tone",
        options=["Negatif", "Netral", "Positif"],
        default=["Negatif", "Netral", "Positif"]
    )
    filter_risiko = st.multiselect(
        "Filter Level Risiko",
        options=["Tinggi", "Sedang", "Rendah"],
        default=["Tinggi", "Sedang", "Rendah"]
    )

    st.markdown("---")
    st.markdown("<div style='font-size:10px;color:#aaa;line-height:1.6'>Analisis Isu Strategis Pengawasan<br>Pusat Strategi Kebijakan Pengawasan BPKP<br>Powered by DeepSeek</div>", unsafe_allow_html=True)


# ── MAIN CONTENT ─────────────────────────────────────────────
# Prioritas sumber data:
# 1. Upload manual via sidebar (override)
# 2. Hasil crawl dari halaman 1 via session_state
# 3. Tidak ada data → tampilkan landing

has_session = st.session_state.get("ais_ready", False) and "hasil" in st.session_state

if uploaded is None and not has_session:
    # Landing state — belum ada data dari mana pun
    st.markdown("""
    <div class="ais-topbar">
      <div>
        <div class="ais-logo">AIS Dashboard</div>
        <div class="ais-subtitle">Analisis Isu Strategis Pengawasan — Pusat Strategi Kebijakan Pengawasan BPKP</div>
      </div>
      <div class="ais-badge">SIAP DIGUNAKAN</div>
    </div>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1,2,1])
    with col2:
        st.markdown("""
        <div style='text-align:center;padding:60px 20px;background:rgba(128,128,128,0.06);border-radius:12px;border:2px dashed rgba(128,128,128,0.3);'>
          <div style='font-size:40px;margin-bottom:12px'>📊</div>
          <div style='font-size:16px;font-weight:600;color:inherit;margin-bottom:8px'>Belum Ada Data</div>
          <div style='font-size:12px;color:inherit;opacity:0.6;line-height:1.6'>
            Jalankan crawl di halaman <b>🔍 Crawl & Analisis</b> terlebih dahulu,<br>
            atau upload file <code>.xlsx</code> hasil crawl sebelumnya<br>
            melalui panel di sebelah kiri.
          </div>
        </div>
        """, unsafe_allow_html=True)
    st.stop()

# ── LOAD & PROCESS DATA ──────────────────────────────────────
if uploaded is not None:
    # Upload manual — selalu override session_state
    df_raw, meta = load_from_excel(uploaded)
    sumber_data = "upload"
    klaster_meta = []  # narasi klaster lengkap tidak tersedia dari Excel, hanya nama per baris

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
    } for i, h in enumerate(hasil_list)])
    meta = {
        "isu":      label_sesi,
        "generate": "dari sesi crawl aktif",
        "total":    str(len(df_raw)),
        "unit":     "Pusat Strategi Kebijakan Pengawasan BPKP"
    }
    sumber_data = "session"

df, stats = compute_stats(df_raw)

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
    <div class="ais-logo">AIS Dashboard</div>
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
tab1, tab2, tab3, tab4 = st.tabs(["📋 Ikhtisar", "🗂️ Daftar Isu", "📈 Tone & Tren", "🔑 Kata Kunci"])


# ════════════════════════════════════════════
# TAB 1 — IKHTISAR
# ════════════════════════════════════════════
with tab1:

    # Stat row
    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        st.markdown(f"""<div class="stat-card" style="border-top:3px solid #F5A623">
          <div class="stat-num">{stats['total']}</div>
          <div class="stat-label">Total Artikel</div>
        </div>""", unsafe_allow_html=True)
    with c2:
        st.markdown(f"""<div class="stat-card" style="border-top:3px solid #E74C3C">
          <div class="stat-num" style="color:#E74C3C">{stats['negatif']}</div>
          <div class="stat-label">Tone Negatif</div>
        </div>""", unsafe_allow_html=True)
    with c3:
        st.markdown(f"""<div class="stat-card" style="border-top:3px solid #95A5A6">
          <div class="stat-num" style="color:#7F8C8D">{stats['netral']}</div>
          <div class="stat-label">Tone Netral</div>
        </div>""", unsafe_allow_html=True)
    with c4:
        st.markdown(f"""<div class="stat-card" style="border-top:3px solid #27AE60">
          <div class="stat-num" style="color:#27AE60">{stats['positif']}</div>
          <div class="stat-label">Tone Positif</div>
        </div>""", unsafe_allow_html=True)
    with c5:
        st.markdown(f"""<div class="stat-card" style="border-top:3px solid #E74C3C">
          <div class="stat-num" style="color:#E74C3C">{stats['pct_neg']}%</div>
          <div class="stat-label">Dominasi Negatif</div>
        </div>""", unsafe_allow_html=True)

    st.markdown("<div style='margin-top:16px'></div>", unsafe_allow_html=True)

    # Spotlight — ambil artikel Negatif dengan teks risiko terpanjang
    # Spotlight — artikel paling relevan dengan topik crawl
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

    if len(df_score) > 0:
        spotlight = df_score.iloc[0]
        tone_spotlight = str(spotlight['Tone'])
        tone_color = {'Negatif': '#E74C3C', 'Netral': '#95A5A6', 'Positif': '#27AE60'}.get(tone_spotlight, '#95A5A6')
        st.markdown(f"""
        <div class="spotlight-box">
          <div class="spotlight-eyebrow">⚡ Isu Prioritas — Paling Relevan & Signifikan</div>
          <div class="spotlight-title">{spotlight['Judul']}</div>
          <div class="spotlight-body">{spotlight['Ringkasan']}<br><br>
            <strong style="color:rgba(255,255,255,0.9)">Risiko:</strong> {spotlight['Risiko']}
          </div>
        </div>
        """, unsafe_allow_html=True)

    # Tone bar + frekuensi subisu
    col_left, col_right = st.columns(2)

    with col_left:
        st.markdown("**Distribusi Tone Pemberitaan**")
        
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

    # Risiko tinggi preview cards
    st.markdown("---")
    st.markdown("**Artikel Risiko Tinggi — Perlu Perhatian**")
    df_tinggi = df[df['level_risiko']=='Tinggi'].head(3)
    if len(df_tinggi) == 0:
        df_tinggi = df[df['Tone']=='Negatif'].head(3)

    if len(df_tinggi) == 0:
        st.info("Tidak ada artikel dengan risiko tinggi atau tone negatif pada periode ini.")
    else:
        cols = st.columns(min(3, len(df_tinggi)))
        for i, (_, row) in enumerate(df_tinggi.iterrows()):
            with cols[i]:
                tone_class = str(row['Tone']).lower()
                aktor_short = str(row['AktorLokasi'])[:40]
                st.markdown(f"""
                <div class="issue-card {tone_class}">
                  <div class="issue-title">{str(row['Judul'])[:80]}{'…' if len(str(row['Judul']))>80 else ''}</div>
                  <div class="issue-sub">{str(row['IsuSubisu'])}</div>
                  <div class="issue-summary">{str(row['Ringkasan'])[:160]}…</div>
                  <div style='margin-top:6px'>
                    <span class="badge badge-{tone_class}">{row['Tone']}</span>
                    <span class="badge badge-aktor">{aktor_short}</span>
                  </div>
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
            # Pilih artikel
            selected_idx = st.session_state.get('selected_idx', 0)

            ada_klaster = 'Klaster' in df_filtered.columns and (df_filtered['Klaster'] != '-').any()

            def render_artikel_item(i, row):
                tone_class = str(row['Tone']).lower()
                is_selected = (i == selected_idx)
                border_style = "border:2px solid rgba(99,179,237,0.8);" if is_selected else "border:1px solid rgba(128,128,128,0.2);"
                bg_style = "background:rgba(99,179,237,0.08);" if is_selected else "background:transparent;"

                judul_short = str(row['Judul'])[:75]+'…' if len(str(row['Judul']))>75 else str(row['Judul'])

                btn_key = f"artikel_{i}"
                st.markdown(f"""
                <div class="issue-card {tone_class}" style="{border_style}{bg_style}">
                  <div style='display:flex;justify-content:space-between;align-items:flex-start;gap:8px'>
                    <div>
                      <div class="issue-title">{judul_short}</div>
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

            if not ada_klaster:
                # Fallback: tampilan datar (Excel lama tanpa kolom Klaster)
                for i, (_, row) in enumerate(df_filtered.iterrows()):
                    render_artikel_item(i, row)
            else:
                # Bangun lookup narasi klaster (jika tersedia dari sesi crawl aktif)
                narasi_klaster = {k.get('nama', '-'): k for k in klaster_meta} if klaster_meta else {}

                # Kelompokkan index df_filtered per nama klaster, urut sesuai urutan klaster_meta
                # jika tersedia, jika tidak urut berdasar jumlah artikel terbanyak dulu.
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

                    # Tone dominan klaster untuk warna aksen expander
                    tone_dom = sub_idx['Tone'].value_counts().idxmax() if jumlah else 'Netral'
                    dom_color = {'Negatif':'#E74C3C','Netral':'#95A5A6','Positif':'#27AE60'}.get(tone_dom,'#95A5A6')

                    label_expander = f"🗂️ {nama}  ·  {jumlah} artikel"
                    with st.expander(label_expander, expanded=(nama == nama_terurut[0])):
                        info_klaster = narasi_klaster.get(nama)
                        if info_klaster:
                            st.markdown(f"""
                            <div style='border-left:3px solid {dom_color};padding:8px 12px;margin-bottom:10px;background:rgba(128,128,128,0.06);border-radius:0 6px 6px 0'>
                              <div style='font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;opacity:0.55;margin-bottom:3px'>Kondisi / Pemicu</div>
                              <div style='font-size:11px;line-height:1.55;margin-bottom:8px'>{info_klaster.get('kondisi_pemicu','-')}</div>
                              <div style='font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;opacity:0.55;margin-bottom:3px'>Akar Persoalan</div>
                              <div style='font-size:11px;line-height:1.55;margin-bottom:8px'>{info_klaster.get('akar_persoalan','-')}</div>
                              <div style='font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;opacity:0.55;margin-bottom:3px'>Risiko Utama</div>
                              <div style='font-size:11px;line-height:1.55;margin-bottom:8px'>{info_klaster.get('risiko_utama','-')}</div>
                              <div style='font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;opacity:0.55;margin-bottom:3px'>Relevansi Pengawasan BPKP</div>
                              <div style='font-size:11px;line-height:1.55'>{info_klaster.get('relevansi_pengawasan','-')}</div>
                            </div>
                            """, unsafe_allow_html=True)

                        for i, row in sub_idx.iterrows():
                            render_artikel_item(i, row)

        with col_detail:
            idx = st.session_state.get('selected_idx', 0)
            if idx < len(df_filtered):
                row = df_filtered.iloc[idx]
                tone_class = str(row['Tone']).lower()

                st.markdown(f"""
                <div id="ais-sticky-detail" style='
                    background:rgba(245,166,35,0.05);
                    border:1px solid rgba(245,166,35,0.35);
                    border-top:4px solid #F5A623;
                    border-radius:10px;
                    padding:20px;
                    box-shadow:0 4px 24px rgba(0,0,0,0.18);
                '>
                  <div style='
                      display:flex;align-items:center;gap:6px;margin-bottom:14px;
                      font-family:"JetBrains Mono",monospace;font-size:10px;
                      letter-spacing:0.08em;color:#F5A623;font-weight:700;
                      text-transform:uppercase;
                  '>
                    📋 Detail Analisis
                  </div>

                  <div style='display:flex;justify-content:space-between;align-items:flex-start;gap:8px;margin-bottom:10px'>
                    <div style='font-size:14px;font-weight:700;color:inherit;line-height:1.4;flex:1'>{row['Judul']}</div>
                    <span class="badge badge-{tone_class}" style='font-size:11px;padding:3px 8px;flex-shrink:0'>{row['Tone']}</span>
                  </div>

                  <div style='display:flex;gap:6px;flex-wrap:wrap;margin-bottom:12px'>
                    <span class="badge badge-aktor">{row['IsuSubisu']}</span>
                    <span class="badge badge-aktor">{row['AktorLokasi']}</span>
                  </div>

                  <hr style='border:none;border-top:1px solid rgba(245,166,35,0.2);margin:10px 0'>

                  <div class="detail-section">
                    <div class="detail-label">Ringkasan Isu</div>
                    <div class="detail-text">{row['Ringkasan']}</div>
                  </div>

                  <div class="detail-section">
                    <div class="detail-label">Risiko</div>
                    <div class="implikasi-box">{row['Risiko']}</div>
                  </div>

                  <div class="detail-section">
                    <div class="detail-label">Area Perhatian</div>
                    <div class="tindaklanjut-box">{row['TindakLanjut']}</div>
                  </div>

                  <hr style='border:none;border-top:1px solid rgba(245,166,35,0.2);margin:10px 0'>
                  <div style='font-size:10px;color:inherit;opacity:0.5'>
                    📅 {row['Tanggal']} &nbsp;·&nbsp;
                    🔗 <a href="{row['Link']}" target="_blank" style='color:#3B82F6'>Buka artikel asli</a>
                  </div>
                </div>
                """, unsafe_allow_html=True)


# ════════════════════════════════════════════
# TAB 3 — TONE & TREN
# ════════════════════════════════════════════
with tab3:
    col_tbl, col_chart = st.columns([3, 2])

    with col_tbl:
        st.markdown("**Tone per Subisu**")
        
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
        st.markdown("**Distribusi Tone Keseluruhan**")
        
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

    # Catatan analitis
    st.markdown("---")
    st.markdown("**Catatan Analitis**")

    neg_issues = df[df['Tone']=='Negatif']['IsuSubisu'].value_counts().head(3).index.tolist()
    pos_issues = df[df['Tone']=='Positif']['IsuSubisu'].value_counts().head(2).index.tolist()

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
