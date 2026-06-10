"""
Media Crawl AIS — Pustrajakwas BPKP
Streamlit web app: input keyword → crawl Google News RSS → analisis Groq → download Excel
"""

import streamlit as st
import feedparser, json, time, re, io
from groq import Groq
from datetime import datetime
from urllib.parse import quote_plus
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Page config ────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Media Crawl AIS — Pustrajakwas",
    page_icon="📰",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── CSS ────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500&display=swap');
html, body, [class*="css"] { font-family: 'IBM Plex Sans', sans-serif !important; }
.stButton>button {
    background: #1F3864; color: white; border: none;
    border-radius: 6px; font-weight: 600; padding: 0.5rem 1.2rem;
}
.stButton>button:hover { background: #2E4D8F; }
.card {
    background: #f8f9fb; border-left: 4px solid #1F3864;
    border-radius: 6px; padding: 1rem 1.2rem; margin-bottom: 0.8rem;
}
.card-neg { border-left-color: #c0392b; }
.card-pos { border-left-color: #27ae60; }
.badge {
    display: inline-block; padding: 2px 10px;
    border-radius: 12px; font-size: 0.75rem; font-weight: 600;
}
.badge-pos { background: #C6EFCE; color: #276221; }
.badge-net { background: #FFEB9C; color: #7D5A00; }
.badge-neg { background: #FFC7CE; color: #9C0006; }
hr.divcard { margin: 0.5rem 0; border-color: #e0e4ed; }
</style>
""", unsafe_allow_html=True)

# ── Tier sumber ────────────────────────────────────────────────────────────
TIER1_KEYWORDS = {
    "kompas", "tempo", "detik", "cnnindonesia", "republika", "antaranews",
    "mediaindonesia", "bisnis", "kontan", "tribunnews", "liputan6", "okezone",
    "sindonews", "jpnn", "suara", "kumparan", "rmol", "inews", "katadata",
    "validnews", "thejakartapost", "jawapos",
}

def tier_sumber(url: str) -> str:
    domain = re.sub(r"https?://(www\.)?", "", url).split("/")[0].lower()
    for t1 in TIER1_KEYWORDS:
        if t1 in domain:
            return "Tier 1"
    return "Tier 2"

# ── Crawl Google News RSS ──────────────────────────────────────────────────
def crawl_google_news(keywords: list, max_articles: int) -> list:
    articles = []
    seen = set()
    headers = {"User-Agent": "Mozilla/5.0 (compatible; AIS-Crawler/1.0)"}

    for kw in keywords:
        query = quote_plus(kw)
        url = f"https://news.google.com/rss/search?q={query}&hl=id&gl=ID&ceid=ID:id"
        try:
            feed = feedparser.parse(url, request_headers=headers)
            for entry in feed.entries:
                link = entry.get("link", "")
                if link in seen:
                    continue
                seen.add(link)

                try:
                    tanggal = datetime(*entry.published_parsed[:3]).strftime("%d %b %Y")
                except Exception:
                    pub = entry.get("published", "")
                    tanggal = pub[:10] if pub else "-"

                tier = tier_sumber(link)
                domain = re.sub(r"https?://(www\.)?", "", link).split("/")[0]
                snippet = re.sub(r"<[^>]+>", "", entry.get("summary", ""))[:500]

                articles.append({
                    "judul":   entry.get("title", "-"),
                    "link":    link,
                    "tanggal": tanggal,
                    "sumber":  domain,
                    "snippet": snippet,
                    "tier":    tier,
                })
                if len(articles) >= max_articles:
                    return articles
        except Exception as e:
            st.warning(f"Gagal crawl '{kw}': {e}")

    return articles

# ── PROMPT SISTEM (v2 — berbasis causal chain) ─────────────────────────────
PROMPT_SISTEM = """Kamu adalah analis isu strategis pengawasan pemerintahan Indonesia, bekerja untuk BPKP Pustrajakwas.

Sebelum menulis analisis, lakukan identifikasi awal:
1. Apa pemicu utama isu ini? Tentukan kategorinya:
   - EKSTERNAL: didorong faktor global (harga komoditas, geopolitik, kebijakan negara lain)
   - KEBIJAKAN: keputusan pemerintah pusat/daerah yang dapat diperdebatkan
   - PELAKSANAAN: kelemahan implementasi program atau penggunaan anggaran
   - TATA KELOLA: indikasi fraud, konflik kepentingan, lemahnya pengendalian internal

2. Siapa yang punya kendali atas isu ini? Apakah pemerintah Indonesia dapat mengubah situasi ini secara langsung, atau hanya merespons?

Gunakan hasil identifikasi itu untuk mengisi JSON berikut (tanpa teks lain di luar JSON):
{
  "ringkasan_isu" : "2-3 kalimat: apa yang terjadi, apa pemicunya (sebutkan eksplisit jika faktor eksternal), dan mengapa relevan bagi pengawasan",
  "isu_subisu"    : "Nama isu utama / subisu spesifik",
  "aktor_lokasi"  : "Nama dan jabatan aktor utama / instansi / lokasi",
  "tone"          : "Positif" atau "Netral" atau "Negatif",
  "risiko_ais"    : "Risiko yang SPESIFIK sesuai kategori isu: jika EKSTERNAL, fokus pada risiko respons kebijakan dan dampak fiskal/sosial; jika PELAKSANAAN atau TATA KELOLA, fokus pada kelemahan pengendalian dan potensi penyimpangan",
  "tindak_lanjut" : "Rekomendasi konkret sesuai kategori: jika isu eksternal, arahkan ke monitoring dampak atau review kesiapan mitigasi; jika tata kelola, arahkan ke audit atau reviu spesifik"
}

Aturan ketat:
- Tone HANYA salah satu dari: Positif, Netral, Negatif
- JANGAN gunakan framing generik seperti "perlu transparansi" atau "perlu keadilan" jika isu dipicu faktor di luar kendali pemerintah
- Bahasa Indonesia formal
- Output HANYA JSON murni, tidak ada teks sebelum atau sesudah, tidak ada markdown fence"""

# ── Analisis per artikel ───────────────────────────────────────────────────
def analisis_artikel(client: Groq, artikel: dict) -> dict:
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
                {"role": "user",   "content": prompt},
            ],
            temperature=0.2,
            max_tokens=700,
        )
        teks = resp.choices[0].message.content.strip()
        teks = re.sub(r"^```json\s*|^```\s*|\s*```$", "", teks).strip()
        hasil = json.loads(teks)
        if hasil.get("tone") not in ("Positif", "Netral", "Negatif"):
            hasil["tone"] = "Netral"
        return hasil
    except Exception:
        return {
            "ringkasan_isu": artikel.get("snippet", "-")[:300],
            "isu_subisu": "-", "aktor_lokasi": "-",
            "tone": "Netral", "risiko_ais": "-", "tindak_lanjut": "-",
        }

# ── Excel builder ──────────────────────────────────────────────────────────
def buat_excel(data: list, label_isu: str) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Identifikasi Isu"

    C_NAVY = "1F3864"; C_WHITE = "FFFFFF"
    C_SUB  = "D9E1F2"; C_ODD   = "EEF2F7"; C_EVEN = "FFFFFF"
    TONE_C = {"Positif": "C6EFCE", "Netral": "FFEB9C", "Negatif": "FFC7CE"}
    NCOL   = 11
    s      = Side(style="thin", color="CCCCCC")
    BD     = Border(left=s, right=s, top=s, bottom=s)

    def style(c, bg, bold=False, sz=9, center=False, fc=C_NAVY):
        c.font      = Font(name="Arial", size=sz, bold=bold, color=fc)
        c.fill      = PatternFill("solid", fgColor=bg)
        c.alignment = Alignment(
            horizontal="center" if center else "left",
            vertical="top", wrap_text=True,
        )
        c.border = BD

    ws.merge_cells(f"A1:{get_column_letter(NCOL)}1")
    c = ws["A1"]
    c.value = "IDENTIFIKASI ISU HARIAN — ANALISIS ISU STRATEGIS PENGAWASAN"
    style(c, C_NAVY, bold=True, sz=13, center=True, fc=C_WHITE)
    ws.row_dimensions[1].height = 28

    ws.merge_cells(f"A2:{get_column_letter(NCOL)}2")
    c = ws["A2"]
    c.value = (
        f"Isu: {label_isu}  |  "
        f"Generate: {datetime.now().strftime('%d %B %Y, %H:%M')}  |  "
        f"Total: {len(data)} artikel  |  Pustrajakwas BPKP"
    )
    style(c, C_SUB, sz=9)
    ws.row_dimensions[2].height = 16
    ws.row_dimensions[3].height = 5

    HEADERS = [
        "No", "Tanggal", "Sumber", "Link/Bukti", "Judul/Post",
        "Ringkasan Isu", "Isu/Subisu", "Aktor/Lokasi",
        "Tone Berita", "Risiko/Implikasi AIS", "Tindak Lanjut",
    ]
    for col, h in enumerate(HEADERS, 1):
        c = ws.cell(row=4, column=col, value=h)
        style(c, C_NAVY, bold=True, sz=10, center=True, fc=C_WHITE)
    ws.row_dimensions[4].height = 34

    for i, d in enumerate(data):
        r  = 5 + i
        bg = C_ODD if i % 2 == 0 else C_EVEN
        baris = [
            i + 1,
            d.get("tanggal", "-"),
            d.get("sumber", "-"),
            d.get("link", "-"),
            d.get("judul", "-"),
            d.get("ringkasan_isu", "-"),
            d.get("isu_subisu", "-"),
            d.get("aktor_lokasi", "-"),
            d.get("tone", "Netral"),
            d.get("risiko_ais", "-"),
            d.get("tindak_lanjut", "-"),
        ]
        for col, val in enumerate(baris, 1):
            c = ws.cell(row=r, column=col, value=val)
            style(c, bg)
        tone_val  = d.get("tone", "Netral")
        tone_cell = ws.cell(row=r, column=9)
        tone_cell.fill = PatternFill("solid", fgColor=TONE_C.get(tone_val, "FFEB9C"))
        ws.row_dimensions[r].height = 60

    col_widths = [5, 12, 18, 35, 40, 45, 25, 25, 12, 45, 40]
    for col, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(col)].width = w

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()

# ── Sidebar ────────────────────────────────────────────────────────────────
with st.sidebar:
    st.title("📰 Media Crawl AIS")
    st.caption("Pustrajakwas BPKP")
    st.divider()

    # API Key — Secrets-aware
    groq_key_default = ""
    if hasattr(st, "secrets"):
        groq_key_default = st.secrets.get("GROQ_API_KEY", "")
    if groq_key_default:
        st.success("✅ API Key terbaca dari Secrets")
        groq_key = groq_key_default
    else:
        groq_key = st.text_input(
            "Groq API Key",
            type="password",
            placeholder="gsk_...",
            help="Daftar gratis di console.groq.com",
        )

    st.divider()
    keywords_raw = st.text_area(
        "Kata Kunci Isu",
        placeholder="Contoh:\nMBG, makan bergizi gratis\nDanantara ekspor\nPertamax BBM",
        height=130,
        help="Satu kata kunci per baris, atau pisahkan dengan koma.",
    )
    label_isu = st.text_input(
        "Label Isu (nama file Excel)",
        placeholder="Contoh: Pertamax BBM Juni 2026",
    )
    max_art = st.slider("Maks. Artikel", min_value=5, max_value=50, value=20, step=5)

    st.divider()
    run_btn = st.button("🔍 Mulai Crawl & Analisis", use_container_width=True)

# ── Main area ──────────────────────────────────────────────────────────────
st.title("Analisis Isu Strategis Pengawasan")
st.caption("Media Crawl otomatis · Analisis Groq (llama-3.3-70b) · Output Excel · Pustrajakwas BPKP")

if run_btn:
    if not groq_key:
        st.error("Masukkan Groq API Key terlebih dahulu.")
        st.stop()
    if not keywords_raw.strip():
        st.error("Masukkan minimal satu kata kunci.")
        st.stop()
    if not label_isu.strip():
        st.error("Isi Label Isu untuk nama file Excel.")
        st.stop()

    keywords = [k.strip() for k in re.split(r"[\n,]+", keywords_raw) if k.strip()]

    st.subheader("⏳ Proses Crawl & Analisis")
    prog_bar  = st.progress(0, text="Memulai crawl...")
    status_tx = st.empty()

    status_tx.info(f"Crawling Google News untuk {len(keywords)} kata kunci...")
    artikel_raw = crawl_google_news(keywords, max_art)

    if not artikel_raw:
        st.warning("Tidak ada artikel ditemukan. Coba kata kunci yang berbeda atau lebih pendek.")
        st.stop()

    status_tx.success(f"✅ {len(artikel_raw)} artikel ditemukan. Memulai analisis Groq...")

    client = Groq(api_key=groq_key)
    hasil_semua = []

    for idx, art in enumerate(artikel_raw):
        pct = int((idx + 1) / len(artikel_raw) * 100)
        prog_bar.progress(pct, text=f"Menganalisis artikel {idx+1}/{len(artikel_raw)}...")
        time.sleep(0.15)

        analisis = analisis_artikel(client, art)
        hasil_semua.append({**art, **analisis})

    prog_bar.progress(100, text="✅ Selesai!")
    status_tx.empty()

    st.session_state["hasil"]     = hasil_semua
    st.session_state["label_isu"] = label_isu.strip()

# ── Tampilkan hasil ────────────────────────────────────────────────────────
if "hasil" in st.session_state:
    hasil_semua = st.session_state["hasil"]
    label_isu   = st.session_state["label_isu"]

    n_pos = sum(1 for h in hasil_semua if h.get("tone") == "Positif")
    n_net = sum(1 for h in hasil_semua if h.get("tone") == "Netral")
    n_neg = sum(1 for h in hasil_semua if h.get("tone") == "Negatif")

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total Artikel", len(hasil_semua))
    col2.metric("🟢 Positif",    n_pos)
    col3.metric("🟡 Netral",     n_net)
    col4.metric("🔴 Negatif",    n_neg)

    excel_bytes = buat_excel(hasil_semua, label_isu)
    fname = f"AIS_{label_isu.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    st.download_button(
        "⬇️ Download Excel",
        data=excel_bytes,
        file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.divider()

    filter_tone = st.selectbox("Filter Tone", ["Semua", "Positif", "Netral", "Negatif"])
    tampil = (
        hasil_semua if filter_tone == "Semua"
        else [h for h in hasil_semua if h.get("tone") == filter_tone]
    )

    st.subheader(f"Pratinjau Artikel ({len(tampil)} ditampilkan)")

    for h in tampil:
        tone      = h.get("tone", "Netral")
        card_cls  = {"Positif": "card-pos", "Negatif": "card-neg"}.get(tone, "")
        badge_cls = {"Positif": "badge-pos", "Negatif": "badge-neg"}.get(tone, "badge-net")
        st.markdown(f"""
        <div class="card {card_cls}">
          <strong>{h.get('judul', '-')}</strong><br>
          <small>{h.get('tanggal', '-')} · {h.get('sumber', '-')} · {h.get('tier', '')}</small>
          &nbsp;<span class="badge {badge_cls}">{tone}</span>
          <hr class="divcard">
          <b>Ringkasan:</b> {h.get('ringkasan_isu', '-')}<br>
          <b>Isu/Subisu:</b> {h.get('isu_subisu', '-')}<br>
          <b>Aktor/Lokasi:</b> {h.get('aktor_lokasi', '-')}<br>
          <b>⚠️ Risiko:</b> {h.get('risiko_ais', '-')}<br>
          <b>✅ Tindak Lanjut:</b> {h.get('tindak_lanjut', '-')}<br>
          <a href="{h.get('link', '#')}" target="_blank"><small>🔗 Buka artikel</small></a>
        </div>
        """, unsafe_allow_html=True)
