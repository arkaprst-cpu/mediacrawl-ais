"""
Media Crawl AIS — Pustrajakwas BPKP
Streamlit web app: input keyword → Groq query expansion → crawl → analisis → download Excel
"""

import streamlit as st
import feedparser, json, time, re, requests, io
from groq import Groq
from datetime import datetime
from urllib.parse import quote_plus
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="Media Crawl AIS — Pustrajakwas",
    page_icon="📰",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500&display=swap');
html, body, [class*="css"] { font-family: 'IBM Plex Sans', sans-serif; }
.main-header { background: linear-gradient(135deg, #1F3864 0%, #2d5299 100%); color: white; padding: 2rem 2.5rem; border-radius: 12px; margin-bottom: 2rem; }
.main-header h1 { font-size: 1.6rem; font-weight: 700; margin: 0 0 0.3rem 0; }
.main-header p { font-size: 0.85rem; opacity: 0.75; margin: 0; font-family: 'IBM Plex Mono', monospace; }
.stat-card { background: white; border: 1px solid #e8ecf0; border-radius: 10px; padding: 1.2rem 1.5rem; text-align: center; box-shadow: 0 2px 8px rgba(0,0,0,0.05); }
.stat-number { font-size: 2rem; font-weight: 700; color: #1F3864; line-height: 1; }
.stat-label { font-size: 0.75rem; color: #6b7280; margin-top: 0.4rem; text-transform: uppercase; letter-spacing: 0.5px; }
.tone-positif { background: #d1fae5; color: #065f46; padding: 2px 10px; border-radius: 20px; font-size: 0.78rem; font-weight: 600; }
.tone-negatif { background: #fee2e2; color: #991b1b; padding: 2px 10px; border-radius: 20px; font-size: 0.78rem; font-weight: 600; }
.tone-netral  { background: #fef3c7; color: #92400e; padding: 2px 10px; border-radius: 20px; font-size: 0.78rem; font-weight: 600; }
.artikel-card { background: white; border: 1px solid #e8ecf0; border-left: 4px solid #1F3864; border-radius: 8px; padding: 1.2rem 1.5rem; margin-bottom: 1rem; box-shadow: 0 1px 4px rgba(0,0,0,0.04); }
.artikel-judul { font-size: 0.95rem; font-weight: 600; color: #1F3864; margin-bottom: 0.4rem; }
.artikel-meta { font-size: 0.78rem; color: #6b7280; font-family: 'IBM Plex Mono', monospace; margin-bottom: 0.8rem; }
.artikel-ringkasan { font-size: 0.85rem; color: #374151; line-height: 1.6; margin-bottom: 0.8rem; }
.artikel-risiko { font-size: 0.82rem; color: #7c3aed; background: #f5f3ff; padding: 0.5rem 0.8rem; border-radius: 6px; margin-bottom: 0.5rem; }
.artikel-tindak { font-size: 0.82rem; color: #065f46; background: #ecfdf5; padding: 0.5rem 0.8rem; border-radius: 6px; }
.query-tag { display: inline-block; background: #e0e7ff; color: #3730a3; padding: 3px 10px; border-radius: 20px; font-size: 0.78rem; margin: 2px; font-family: 'IBM Plex Mono', monospace; }
section[data-testid="stSidebar"] { background: #f8fafc; border-right: 1px solid #e8ecf0; }
.stButton > button { background: #1F3864 !important; color: white !important; border: none !important; border-radius: 8px !important; font-weight: 600 !important; width: 100%; }
.stButton > button:hover { background: #2d5299 !important; }
</style>
""", unsafe_allow_html=True)

RSS_FEEDS = [
    ("Antara",         "Tier 1", "https://www.antaranews.com/rss/terkini"),
    ("Antara Ekonomi", "Tier 1", "https://www.antaranews.com/rss/ekonomi"),
    ("Antara Politik", "Tier 1", "https://www.antaranews.com/rss/politik"),
    ("Tempo",          "Tier 1", "https://rss.tempo.co/"),
    ("Tempo Bisnis",   "Tier 1", "https://rss.tempo.co/bisnis"),
    ("Tempo Nasional", "Tier 1", "https://rss.tempo.co/nasional"),
    ("Detik",          "Tier 1", "https://rss.detik.com/index.php/detikcom"),
    ("CNBC Indonesia", "Tier 1", "https://www.cnbcindonesia.com/rss"),
    ("CNN Indonesia",  "Tier 1", "https://www.cnnindonesia.com/rss"),
    ("Kompas",         "Tier 1", "https://indeks.kompas.com/terbaru/xml/rss20.xml"),
    ("Republika",      "Tier 1", "https://rss.republika.co.id/rss/nasional"),
    ("Bisnis",         "Tier 1", "https://bisnis.com/rss"),
    ("Kontan",         "Tier 2", "https://rss.kontan.co.id/rss/nasional"),
    ("Katadata",       "Tier 2", "https://katadata.co.id/rss.xml"),
    ("Tirto",          "Tier 2", "https://tirto.id/rss"),
    ("JPNN",           "Tier 2", "https://www.jpnn.com/rss/terbaru"),
    ("Mediaindonesia", "Tier 2", "https://mediaindonesia.com/rss/terbaru"),
    ("Okezone",        "Tier 2", "https://rss.okezone.com/ekonomi"),
    ("Suara.com",      "Tier 2", "https://rss.suara.com/news.rss"),
]

PROMPT_SISTEM = """Kamu adalah analis isu strategis pengawasan pemerintahan Indonesia.
Sudut pandang: tata kelola, risiko kebijakan, akuntabilitas, dan pengawasan intern.

Untuk setiap berita, hasilkan JSON berikut (tanpa teks lain di luar JSON):
{
  "ringkasan_isu"   : "2-3 kalimat: apa yang terjadi dan mengapa penting dari sisi pengawasan",
  "isu_subisu"      : "Nama isu utama / subisu spesifik",
  "aktor_lokasi"    : "Nama dan jabatan aktor utama / instansi / lokasi",
  "tone"            : "Positif" atau "Netral" atau "Negatif",
  "risiko_ais"      : "Risiko tata kelola atau implikasi pengawasan yang relevan bagi BPKP",
  "tindak_lanjut"   : "Satu rekomendasi konkret untuk pengawasan atau kajian BPKP"
}
Aturan: Tone HANYA Positif/Netral/Negatif. Bahasa Indonesia formal. Output HANYA JSON murni."""

HEADERS_HTTP = {"User-Agent": "Mozilla/5.0 (compatible; MediaCrawlBot/1.0)"}

# ── Helper functions ──────────────────────────────────────────
def parse_tanggal(entry):
    for field in ["published", "updated"]:
        val = entry.get(field, "")
        if val:
            try:
                from email.utils import parsedate_to_datetime
                return parsedate_to_datetime(val).strftime("%d %b %Y")
            except Exception:
                return val[:10]
    return "-"

def bersihkan_html(teks):
    return re.sub(r"<[^>]+>", " ", teks).strip()

def ekspansi_keyword(client, keyword_input):
    prompt = f"""Kamu adalah asisten pencarian berita media Indonesia.

User ingin mencari berita terkait: "{keyword_input}"

Tugasmu: hasilkan 5 query pencarian yang BERBEDA-BEDA untuk Google News,
agar berita yang ditemukan lebih lengkap dan kontekstual.

Aturan:
- Variasikan: singkatan, nama lengkap, sinonim, aktor kunci, lokasi jika ada
- Setiap query pendek (2-5 kata), seperti yang orang ketik di search engine
- Gunakan bahasa Indonesia
- Output HANYA array JSON, tidak ada teks lain

Contoh untuk input "MBG Jawa Tengah":
["MBG Jawa Tengah", "Makan Bergizi Gratis Semarang", "program makan siang sekolah Jateng", "dapur MBG Solo Semarang", "Makan Bergizi Gratis Jateng evaluasi"]

Sekarang buat untuk: "{keyword_input}"
"""
    try:
        resp = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.4, max_tokens=300
        )
        teks = resp.choices[0].message.content.strip()
        teks = re.sub(r"^```json\s*|^```\s*|\s*```$", "", teks).strip()
        queries = json.loads(teks)
        if isinstance(queries, list) and len(queries) > 0:
            return [str(q) for q in queries[:6]]
    except Exception:
        pass
    return [k.strip() for k in keyword_input.split(",") if k.strip()]

def crawl_google_news(queries, max_hasil=50):
    hasil, link_sudah = [], set()
    for q in queries:
        url = f"https://news.google.com/rss/search?q={quote_plus(q)}&hl=id&gl=ID&ceid=ID:id"
        try:
            resp = requests.get(url, headers=HEADERS_HTTP, timeout=15)
            feed = feedparser.parse(resp.text)
            for entry in feed.entries:
                link = entry.get("link", "")
                if link in link_sudah: continue
                link_sudah.add(link)
                judul_raw = entry.get("title", "")
                summary   = bersihkan_html(entry.get("summary", ""))
                bagian    = judul_raw.rsplit(" - ", 1)
                hasil.append({
                    "tanggal": parse_tanggal(entry),
                    "sumber" : bagian[-1].strip() if len(bagian) > 1 else "Google News",
                    "link"   : link,
                    "judul"  : bagian[0].strip() if len(bagian) > 1 else judul_raw,
                    "snippet": summary[:600] if summary else judul_raw,
                    "tier"   : "Tier 1",
                })
        except Exception:
            pass
        time.sleep(0.5)
    return hasil[:max_hasil]

def crawl_rss_feeds(kata_kunci_asli):
    hasil = []
    for nama, tier, url in RSS_FEEDS:
        try:
            resp = requests.get(url, headers=HEADERS_HTTP, timeout=10)
            resp.raise_for_status()
            feed = feedparser.parse(resp.text)
            for entry in feed.entries:
                judul   = entry.get("title", "")
                summary = bersihkan_html(entry.get("summary", ""))
                if not any(k.lower() in (judul+summary).lower() for k in kata_kunci_asli): continue
                hasil.append({
                    "tanggal": parse_tanggal(entry),
                    "sumber" : nama,
                    "link"   : entry.get("link", ""),
                    "judul"  : judul,
                    "snippet": summary[:600],
                    "tier"   : tier,
                })
        except Exception:
            pass
        time.sleep(0.3)
    return hasil

def analisis_groq(client, artikel):
    prompt = (
        f"Judul   : {artikel['judul']}\n"
        f"Sumber  : {artikel['sumber']}\n"
        f"Tanggal : {artikel['tanggal']}\n"
        f"Konten  : {artikel['snippet']}\n\nHasilkan JSON analisis."
    )
    try:
        resp = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[
                {"role": "system", "content": PROMPT_SISTEM},
                {"role": "user",   "content": prompt}
            ],
            temperature=0.2, max_tokens=600
        )
        teks = resp.choices[0].message.content.strip()
        teks = re.sub(r"^```json\s*|^```\s*|\s*```$", "", teks).strip()
        hasil = json.loads(teks)
        if hasil.get("tone") not in ("Positif", "Netral", "Negatif"):
            hasil["tone"] = "Netral"
        return hasil
    except Exception:
        return {"ringkasan_isu": artikel.get("snippet","-")[:300],
                "isu_subisu": "-", "aktor_lokasi": "-",
                "tone": "Netral", "risiko_ais": "-", "tindak_lanjut": "-"}

def buat_excel(data, label_isu):
    wb = Workbook(); ws = wb.active; ws.title = "Identifikasi Isu"
    C_NAVY="1F3864"; C_WHITE="FFFFFF"; C_SUB="D9E1F2"; C_ODD="EEF2F7"; C_EVEN="FFFFFF"
    TIER_C={"Tier 1":"C6EFCE","Tier 2":"FFEB9C"}
    TONE_C={"Positif":"C6EFCE","Netral":"FFEB9C","Negatif":"FFC7CE"}
    NCOL=11; s=Side(style="thin",color="CCCCCC"); BD=Border(left=s,right=s,top=s,bottom=s)
    def style(c,bg,bold=False,sz=9,center=False,fc=C_NAVY):
        c.font=Font(name="Arial",size=sz,bold=bold,color=fc)
        c.fill=PatternFill("solid",fgColor=bg)
        c.alignment=Alignment(horizontal="center" if center else "left",vertical="top",wrap_text=True)
        c.border=BD
    ws.merge_cells(f"A1:{get_column_letter(NCOL)}1")
    c=ws["A1"]; c.value="IDENTIFIKASI ISU HARIAN — ANALISIS ISU STRATEGIS PENGAWASAN"
    style(c,C_NAVY,bold=True,sz=13,center=True,fc=C_WHITE); ws.row_dimensions[1].height=28
    ws.merge_cells(f"A2:{get_column_letter(NCOL)}2")
    c=ws["A2"]; c.value=f"Isu: {label_isu}  |  Generate: {datetime.now().strftime('%d %B %Y, %H:%M')}  |  Total: {len(data)} artikel  |  Pustrajakwas BPKP"
    style(c,C_SUB,sz=9); ws.row_dimensions[2].height=16; ws.row_dimensions[3].height=5
    HEADERS=["No","Tanggal","Sumber","Link/Bukti","Judul/Post","Ringkasan Isu","Isu/Subisu","Aktor/Lokasi","Tone Berita","Risiko/Implikasi AIS","Tindak Lanjut"]
    for col,h in enumerate(HEADERS,1):
        c=ws.cell(row=4,column=col,value=h); style(c,C_NAVY,bold=True,sz=10,center=True,fc=C_WHITE)
    ws.row_dimensions[4].height=34
    for i,d in enumerate(data):
        r=5+i; bg=C_ODD if i%2==0 else C_EVEN
        baris=[i+1,d.get("tanggal","-"),d.get("sumber","-"),d.get("link","-"),d.get("judul","-"),
               d.get("ringkasan_isu","-"),d.get("isu_subisu","-"),d.get("aktor_lokasi","-"),
               d.get("tone","Netral"),d.get("risiko_ais","-"),d.get("tindak_lanjut","-")]
        for col,val in enumerate(baris,1): style(ws.cell(row=r,column=col,value=val),bg)
        c_src=ws.cell(row=r,column=3); c_src.fill=PatternFill("solid",fgColor=TIER_C.get(d.get("tier","Tier 2"),"FCE4D6")); c_src.font=Font(name="Arial",size=9,bold=True)
        tone=d.get("tone","Netral"); c_ton=ws.cell(row=r,column=9)
        c_ton.fill=PatternFill("solid",fgColor=TONE_C.get(tone,"FFEB9C")); c_ton.font=Font(name="Arial",size=9,bold=True)
        c_ton.alignment=Alignment(horizontal="center",vertical="top",wrap_text=True); ws.row_dimensions[r].height=75
    for i,w in enumerate([5,12,16,32,36,44,26,28,13,44,38],1): ws.column_dimensions[get_column_letter(i)].width=w
    fr=5+len(data)+1; ws.merge_cells(f"A{fr}:{get_column_letter(NCOL)}{fr}")
    c=ws[f"A{fr}"]; c.value="Tier 1=Nasional (hijau) | Tier 2=Regional (kuning) | Tone: Hijau=Positif | Kuning=Netral | Merah=Negatif | Analisis: Groq Llama 3.3 70B"
    c.font=Font(name="Arial",italic=True,size=8,color="595959"); c.fill=PatternFill("solid",fgColor="F2F2F2")
    c.alignment=Alignment(horizontal="left",vertical="center",wrap_text=True); ws.row_dimensions[fr].height=20
    ws.freeze_panes="A5"; buf=io.BytesIO(); wb.save(buf); buf.seek(0); return buf

def render_hasil():
    """Render semua hasil dari session_state — dipanggil setelah crawl selesai maupun saat re-run."""
    hasil_list      = st.session_state.hasil_list
    queries_expanded = st.session_state.queries_expanded
    keyword_input   = st.session_state.keyword_input
    label_isu       = st.session_state.label_isu
    gn_count        = st.session_state.gn_count
    rss_count       = st.session_state.rss_count

    # Query expansion info
    tags_html = " ".join([f'<span class="query-tag">🔍 {q}</span>' for q in queries_expanded])
    st.markdown(f"""
    <div style="background:#f0f4ff;border:1px solid #c7d2fe;border-radius:8px;padding:1rem 1.2rem;margin-bottom:1rem;">
        <div style="font-size:0.78rem;color:#4338ca;font-weight:600;margin-bottom:0.5rem">
            ✨ AI mengubah "{keyword_input}" menjadi {len(queries_expanded)} query pencarian:
        </div>
        {tags_html}
    </div>
    """, unsafe_allow_html=True)

    # Statistik crawl
    st.markdown("### 📡 Hasil Crawl")
    c1,c2,c3 = st.columns(3)
    for col, num, label in [(c1, len(hasil_list), "Total Artikel"), (c2, gn_count, "Google News"), (c3, rss_count, "RSS Media")]:
        with col:
            st.markdown(f'<div class="stat-card"><div class="stat-number">{num}</div><div class="stat-label">{label}</div></div>', unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)

    # Statistik tone
    tone_counts = {"Positif":0,"Netral":0,"Negatif":0}
    for h in hasil_list: tone_counts[h.get("tone","Netral")] = tone_counts.get(h.get("tone","Netral"),0)+1
    st.markdown("### 📊 Ringkasan Tone Berita")
    c1,c2,c3 = st.columns(3)
    for col, tone, warna, emoji in [(c1,"Positif","#065f46","🟢"),(c2,"Netral","#92400e","🟡"),(c3,"Negatif","#991b1b","🔴")]:
        with col:
            st.markdown(f'<div class="stat-card"><div class="stat-number" style="color:{warna}">{tone_counts[tone]}</div><div class="stat-label">{emoji} {tone}</div></div>', unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)

    # Download
    st.markdown("### ⬇️ Unduh Hasil")
    excel_buf = buat_excel(hasil_list, label_isu)
    nama_file = f"MediaCrawl_AIS_{label_isu.replace(' ','_')}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    st.download_button(
        "📥 Download Excel", data=excel_buf, file_name=nama_file,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
    st.markdown("<br>", unsafe_allow_html=True)

    # Pratinjau
    st.markdown("### 📋 Pratinjau Hasil Analisis")
    filter_tone = st.selectbox("Filter tone:", ["Semua","Positif","Netral","Negatif"])
    tampil = hasil_list if filter_tone=="Semua" else [h for h in hasil_list if h.get("tone")==filter_tone]
    for h in tampil:
        tone = h.get("tone","Netral")
        st.markdown(f"""
        <div class="artikel-card">
            <div class="artikel-judul">{h.get('judul','-')}</div>
            <div class="artikel-meta">📅 {h.get('tanggal','-')} &nbsp;·&nbsp; 📰 {h.get('sumber','-')} &nbsp;·&nbsp; <span class="tone-{tone.lower()}">{tone}</span></div>
            <div class="artikel-ringkasan">{h.get('ringkasan_isu','-')}</div>
            <div class="artikel-risiko">⚠️ <b>Risiko:</b> {h.get('risiko_ais','-')}</div>
            <div class="artikel-tindak">✅ <b>Tindak Lanjut:</b> {h.get('tindak_lanjut','-')}</div>
        </div>""", unsafe_allow_html=True)
        with st.expander("🔗 Lihat link artikel"):
            st.write(h.get("link","-"))

# ══════════════════════════════════════════════════════════════
# UI
# ══════════════════════════════════════════════════════════════
st.markdown("""
<div class="main-header">
    <h1>📰 Media Crawl — Analisis Isu Strategis</h1>
    <p>Pusat Strategi Kebijakan Pengawasan · Pustrajakwas BPKP</p>
</div>
""", unsafe_allow_html=True)

# ── Inisialisasi session_state ────────────────────────────────
if "hasil_list" not in st.session_state:
    st.session_state.hasil_list       = []
    st.session_state.queries_expanded = []
    st.session_state.keyword_input    = ""
    st.session_state.label_isu        = ""
    st.session_state.gn_count         = 0
    st.session_state.rss_count        = 0

with st.sidebar:
    st.markdown("### ⚙️ Konfigurasi")
    groq_key_secret = st.secrets.get("GROQ_API_KEY", "") if hasattr(st, "secrets") else ""
    if groq_key_secret:
        st.success("✅ API Key terbaca dari Secrets", icon="🔑")
        groq_key = groq_key_secret
    else:
        groq_key = st.text_input("Groq API Key", type="password", placeholder="gsk_...", help="Daftar gratis di console.groq.com")
    st.markdown("---")
    st.markdown("### 🔍 Parameter Crawl")
    keyword_input = st.text_input("Kata Kunci Isu", placeholder="Contoh: MBG Jawa Tengah", help="Tulis natural — AI akan perluas otomatis")
    label_isu     = st.text_input("Label Isu (untuk nama file)", placeholder="Contoh: MBG_Jateng")
    max_artikel   = st.slider("Maks. Artikel Diproses", min_value=5, max_value=30, value=20, step=5)
    st.markdown("---")
    tombol_crawl = st.button("🚀 Mulai Crawl & Analisis", use_container_width=True)
    st.markdown("---")
    st.markdown("""<div style="font-size:0.75rem;color:#6b7280;line-height:1.8">
    <b>Cara pakai:</b><br>
    1. Ketik isu secara natural<br>
    &nbsp;&nbsp;&nbsp;<i>Contoh: "MBG Jawa Tengah"</i><br>
    2. AI otomatis perluas keyword<br>
    3. Klik Mulai Crawl<br>
    4. Unduh hasil Excel
    </div>""", unsafe_allow_html=True)

# ── Area utama ────────────────────────────────────────────────
if tombol_crawl:
    # Validasi
    if not groq_key:
        st.error("❌ Groq API Key belum diisi."); st.stop()
    if not keyword_input.strip():
        st.error("❌ Kata kunci belum diisi."); st.stop()
    if not label_isu.strip():
        label_isu = keyword_input.split(",")[0].strip().replace(" ", "_")
    kata_kunci_asli = [k.strip() for k in keyword_input.replace(",", " ").split() if k.strip()]

    try:
        client = Groq(api_key=groq_key)
        client.chat.completions.create(model="llama-3.3-70b-versatile",
            messages=[{"role":"user","content":"OK"}], max_tokens=3)
    except Exception as e:
        st.error(f"❌ Groq API gagal: {e}"); st.stop()

    # Tahap 0: Query expansion
    st.markdown("### 🧠 Tahap 0 — Ekspansi Keyword oleh AI")
    with st.spinner("AI sedang memperluas keyword..."):
        queries_expanded = ekspansi_keyword(client, keyword_input)

    # Tahap 1: Crawl
    st.markdown("### 📡 Tahap 1 — Crawl Media")
    status_crawl = st.empty(); prog_crawl = st.progress(0)
    status_crawl.info("🌐 Mengambil artikel dari Google News...")
    semua, link_set = [], set()
    gn = crawl_google_news(queries_expanded, max_hasil=50)
    for a in gn:
        if a["link"] not in link_set: semua.append(a); link_set.add(a["link"])
    prog_crawl.progress(40)
    status_crawl.info("📡 Mengambil dari RSS media Indonesia...")
    rss = crawl_rss_feeds(kata_kunci_asli)
    for a in rss:
        if a["link"] not in link_set: semua.append(a); link_set.add(a["link"])
    prog_crawl.progress(100)
    semua = semua[:max_artikel]
    if not semua:
        st.warning("⚠️ Tidak ada artikel. Coba kata kunci berbeda."); st.stop()
    status_crawl.success(f"✅ Ditemukan **{len(semua)} artikel** relevan.")

    # Tahap 2: Analisis
    st.markdown("### 🤖 Tahap 2 — Analisis dengan Groq AI")
    status_ai = st.empty(); prog_ai = st.progress(0); hasil_list = []
    for i, artikel in enumerate(semua):
        status_ai.info(f"Menganalisis [{i+1}/{len(semua)}]: *{artikel['judul'][:70]}*")
        hasil_list.append({**artikel, **analisis_groq(client, artikel)})
        prog_ai.progress(int((i+1)/len(semua)*100))
        time.sleep(2.5)
    prog_ai.progress(100)
    status_ai.success(f"✅ Selesai — {len(hasil_list)} artikel dianalisis.")

    # ── Simpan ke session_state ───────────────────────────────
    st.session_state.hasil_list       = hasil_list
    st.session_state.queries_expanded = queries_expanded
    st.session_state.keyword_input    = keyword_input
    st.session_state.label_isu        = label_isu
    st.session_state.gn_count         = len(gn)
    st.session_state.rss_count        = len(rss)

# ── Tampilkan hasil (dari session_state) ─────────────────────
if st.session_state.hasil_list:
    render_hasil()
else:
    st.markdown("""
    <div style="text-align:center;padding:4rem 2rem;color:#9ca3af;">
        <div style="font-size:3rem;margin-bottom:1rem">🗞️</div>
        <div style="font-size:1.1rem;font-weight:600;color:#374151;margin-bottom:0.5rem">Siap melakukan crawl</div>
        <div style="font-size:0.88rem">Ketik isu di panel kiri — AI akan otomatis memperluas pencarian</div>
    </div>""", unsafe_allow_html=True)
