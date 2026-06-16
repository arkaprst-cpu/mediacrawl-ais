"""
Media Crawl AIS — Pusat Strategi Kebijakan Pengawasan BPKP
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
    page_title="Media Crawl AIS — Pusat Strategi Kebijakan Pengawasan",
    page_icon="📰",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── NAVIGASI HALAMAN ───────────────────────────────────────────────────────
# Ini satu-satunya blok sidebar untuk navigasi. Sidebar konten crawl
# ada di dalam blok if page == "🔍 Crawl & Analisis" di bawah.
with st.sidebar:
    st.markdown("""
    <div style='background:linear-gradient(135deg,#0D1B2A,#1C3D5A);
                border-radius:8px;padding:12px 16px;margin-bottom:12px;
                border-bottom:2px solid #F5A623'>
      <div style='font-family:monospace;font-size:15px;font-weight:700;color:#F5A623'>AIS</div>
      <div style='font-size:10px;color:rgba(255,255,255,0.6);margin-top:2px'>Pusat Strategi Kebijakan Pengawasan BPKP</div>
    </div>
    """, unsafe_allow_html=True)

    page = st.radio(
        "Navigasi",
        options=["🔍 Crawl & Analisis", "📊 Dashboard AIS"],
        label_visibility="collapsed"
    )
    st.divider()

# ── ROUTING ────────────────────────────────────────────────────────────────

# ══════════════════════════════════════════════════════════════════════════
# HALAMAN 1 — CRAWL & ANALISIS
# ══════════════════════════════════════════════════════════════════════════
if page == "🔍 Crawl & Analisis":

    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500&display=swap');
    html, body, [class*="css"] { font-family: 'IBM Plex Sans', sans-serif; }

    .main-header {
        background: linear-gradient(135deg, #1F3864 0%, #2d5299 100%);
        color: white; padding: 2rem 2.5rem; border-radius: 12px; margin-bottom: 2rem;
    }
    .main-header h1 { font-size: 1.6rem; font-weight: 700; margin: 0 0 0.3rem 0; }
    .main-header p  { font-size: 0.85rem; opacity: 0.75; margin: 0; font-family: 'IBM Plex Mono', monospace; }

    .stat-card {
        background: white; border: 1px solid #e8ecf0; border-radius: 10px;
        padding: 1.2rem 1.5rem; text-align: center;
        box-shadow: 0 2px 8px rgba(0,0,0,0.05);
    }
    .stat-number { font-size: 2rem; font-weight: 700; color: #1F3864; line-height: 1; }
    .stat-label  { font-size: 0.75rem; color: #6b7280; margin-top: 0.4rem; text-transform: uppercase; letter-spacing: 0.05em; }

    .artikel-card {
        background: #ffffff; border: 1px solid #e2e8f0;
        border-left: 4px solid #1F3864; border-radius: 8px;
        padding: 1.2rem 1.5rem; margin-bottom: 1rem;
        box-shadow: 0 1px 4px rgba(0,0,0,0.06);
    }
    .artikel-card:hover { box-shadow: 0 4px 12px rgba(31,56,100,0.12); }
    .artikel-judul {
        font-size: 1rem; font-weight: 600; color: #1e293b;
        margin-bottom: 0.4rem; line-height: 1.4;
    }
    .artikel-meta {
        font-size: 0.78rem; color: #64748b; margin-bottom: 0.8rem;
        font-family: 'IBM Plex Mono', monospace;
    }
    .artikel-ringkasan {
        font-size: 0.88rem; color: #374151; line-height: 1.6;
        margin-bottom: 0.6rem;
    }
    .artikel-risiko {
        font-size: 0.85rem; color: #7c3aed; background: #f5f3ff;
        border-left: 3px solid #7c3aed; padding: 0.5rem 0.8rem;
        border-radius: 0 6px 6px 0; margin-bottom: 0.5rem;
    }
    .artikel-tindak {
        font-size: 0.85rem; color: #065f46; background: #ecfdf5;
        border-left: 3px solid #059669; padding: 0.5rem 0.8rem;
        border-radius: 0 6px 6px 0;
    }

    .tone-positif {
        background: #d1fae5; color: #065f46;
        padding: 2px 10px; border-radius: 12px;
        font-size: 0.75rem; font-weight: 600;
    }
    .tone-netral {
        background: #fef3c7; color: #92400e;
        padding: 2px 10px; border-radius: 12px;
        font-size: 0.75rem; font-weight: 600;
    }
    .tone-negatif {
        background: #fee2e2; color: #991b1b;
        padding: 2px 10px; border-radius: 12px;
        font-size: 0.75rem; font-weight: 600;
    }

    .query-box {
        background: #eff6ff; border: 1px solid #bfdbfe;
        border-radius: 8px; padding: 0.8rem 1rem; margin-bottom: 1rem;
        font-size: 0.83rem; color: #1e40af;
        font-family: 'IBM Plex Mono', monospace;
    }
    </style>
    """, unsafe_allow_html=True)

    # ── Tier sumber ────────────────────────────────────────────────────────
    TIER1_KEYWORDS = {
        "kompas","tempo","detik","cnnindonesia","republika","antaranews",
        "mediaindonesia","bisnis","kontan","tribunnews","liputan6","okezone",
        "sindonews","jpnn","suara","kumparan","rmol","inews","katadata",
        "validnews","thejakartapost","jawapos",
    }

    def tier_sumber(url: str) -> str:
        domain = re.sub(r"https?://(www\.)?", "", url).split("/")[0].lower()
        for t1 in TIER1_KEYWORDS:
            if t1 in domain:
                return "Tier 1"
        return "Tier 2"

    def extract_sumber_dari_judul(judul: str) -> str:
        """
        Google News RSS format judul: 'Judul Artikel - NamaMedia'
        Ekstrak NamaMedia dari bagian akhir judul.
        Lebih reliable daripada resolve redirect Google News.
        """
        m = re.search(r"\s[-\u2013]\s([^-\u2013]+)$", judul.strip())
        if m:
            sumber = m.group(1).strip()
            sumber = re.sub(r"\s*\[.*?\]\s*$", "", sumber).strip()
            return sumber if sumber else ""
        return ""

    # ── Query expansion ────────────────────────────────────────────────────
    def ekspansi_keyword(client: Groq, keyword: str) -> list:
        prompt = f"""Kamu adalah asisten pencarian berita. Dari input keyword berikut, buat 4-5 variasi query pencarian berita yang lebih spesifik dan efektif untuk Google News.

    Keyword input: "{keyword}"

    Aturan:
    - Variasikan dengan sinonim, singkatan, nama lokasi spesifik, atau aspek berbeda dari isu yang sama
    - Gunakan bahasa Indonesia
    - Kembalikan HANYA array JSON berisi string query, tanpa teks lain

    Contoh output: ["query 1", "query 2", "query 3", "query 4"]"""
        try:
            resp = client.chat.completions.create(
                model="llama-3.3-70b-versatile",
                messages=[{"role": "user", "content": prompt}],
                temperature=0.4,
                max_tokens=300,
            )
            teks = resp.choices[0].message.content.strip()
            teks = re.sub(r"^```json\s*|^```\s*|\s*```$", "", teks).strip()
            queries = json.loads(teks)
            if isinstance(queries, list):
                return [keyword] + [q for q in queries if isinstance(q, str)]
        except Exception:
            pass
        return [keyword]

    # ── Filter konten video ───────────────────────────────────────────────
    DOMAIN_VIDEO = {
        "kompas.tv", "metrotvnews.com", "tvone.co.id",
        "rctiplus.com", "vidio.com", "youtube.com", "youtu.be",
    }
    JUDUL_VIDEO = [
        "[full]", "[live]", "[video]", "[breaking]",
        "live streaming", "siaran langsung", "tonton video",
        "nonton:", "breaking news:", "full video",
    ]

    def is_video(judul: str, domain: str) -> bool:
        judul_lower = judul.lower()
        for dv in DOMAIN_VIDEO:
            if dv in domain.lower():
                return True
        for marker in JUDUL_VIDEO:
            if marker in judul_lower:
                return True
        return False

    # ── Bersihkan snippet RSS ─────────────────────────────────────────────
    def bersihkan_snippet(snippet: str, judul: str) -> str:
        """
        Bersihkan snippet RSS dari HTML entities dan pengulangan judul.
        Google News RSS hanya memberikan judul + nama media, bukan konten.
        Return: string bersih, atau "" jika isinya hanya judul.
        """
        # Decode HTML entities
        teks = snippet.replace("&nbsp;", " ").replace("&amp;", "&")
        teks = re.sub(r"&[a-z]+;", " ", teks)
        teks = re.sub(r"<[^>]+>", "", teks)
        teks = re.sub(r"\s+", " ", teks).strip()

        # Buang jika isinya hanya pengulangan judul ± nama media
        judul_bersih = re.sub(r"\s+", " ", judul).strip().lower()
        teks_lower = teks.lower()
        # Cek apakah teks dimulai dengan judul (pattern Google News RSS)
        if teks_lower.startswith(judul_bersih[:40].lower()):
            return ""
        return teks

    # ── Crawl Google News RSS ──────────────────────────────────────────────
    def crawl_google_news(queries: list, max_articles: int) -> list:
        articles = []
        seen = set()
        skipped = 0
        headers = {"User-Agent": "Mozilla/5.0 (compatible; AIS-Crawler/1.0)"}

        for q in queries:
            url = f"https://news.google.com/rss/search?q={quote_plus(q)}&hl=id&gl=ID&ceid=ID:id"
            try:
                feed = feedparser.parse(url, request_headers=headers)
                for entry in feed.entries:
                    link = entry.get("link", "")
                    if link in seen:
                        continue
                    seen.add(link)

                    judul  = entry.get("title", "-")
                    domain = re.sub(r"https?://(www\.)?", "", link).split("/")[0]

                    # Filter artikel video
                    if is_video(judul, domain):
                        skipped += 1
                        continue

                    try:
                        tanggal = datetime(*entry.published_parsed[:3]).strftime("%d %b %Y")
                    except Exception:
                        pub = entry.get("published", "")
                        tanggal = pub[:10] if pub else "-"

                    # Bersihkan snippet RSS
                    snippet_raw = entry.get("summary", "")
                    snippet_rss = re.sub(r"<[^>]+>", "", snippet_raw)
                    konten = bersihkan_snippet(snippet_rss, judul)
                    # konten bisa kosong — itu normal, Groq akan analisis dari judul

                    # Ekstrak nama sumber dari judul (paling reliable untuk Google News)
                    sumber_nama = extract_sumber_dari_judul(judul)
                    if not sumber_nama:
                        sumber_nama = re.sub(r"https?://(www\.)?", "", link).split("/")[0]
                    tier_asli = tier_sumber(sumber_nama.lower())

                    articles.append({
                        "judul":   judul,
                        "link":    link,
                        "tanggal": tanggal,
                        "sumber":  sumber_nama,
                        "snippet": konten,
                        "tier":    tier_asli,
                    })
                    if len(articles) >= max_articles:
                        if skipped > 0:
                            st.caption(f"ℹ️ {skipped} artikel video/konten kosong dilewati.")
                        return articles
            except Exception as e:
                st.warning(f"Gagal crawl '{q}': {e}")

        if skipped > 0:
            st.caption(f"ℹ️ {skipped} artikel video/konten kosong dilewati.")
        return articles

    # ── PROMPT SISTEM ──────────────────────────────────────────────────────
    PROMPT_SISTEM = """Kamu adalah analis isu strategis pengawasan pemerintahan Indonesia, bekerja untuk BPKP Pusat Strategi Kebijakan Pengawasan.

    Sebelum menulis analisis, lakukan identifikasi awal:
    1. Apa pemicu utama isu ini? Tentukan kategorinya:
       - EKSTERNAL: didorong faktor global (harga komoditas, geopolitik, kebijakan negara lain)
       - KEBIJAKAN: keputusan pemerintah pusat/daerah yang dapat diperdebatkan
       - PELAKSANAAN: kelemahan implementasi program atau penggunaan anggaran
       - TATA KELOLA: indikasi fraud, konflik kepentingan, lemahnya pengendalian internal

    2. Siapa yang punya kendali atas isu ini? Apakah pemerintah Indonesia dapat mengubah situasi ini secara langsung, atau hanya merespons?

    Gunakan hasil identifikasi itu untuk mengisi JSON berikut (tanpa teks lain di luar JSON):
    {
      "ringkasan_isu"  : "2-3 kalimat: apa yang terjadi, apa pemicunya (sebutkan eksplisit jika faktor eksternal), dan mengapa relevan bagi pengawasan",
      "isu_subisu"     : "Nama isu utama / subisu spesifik",
      "aktor_lokasi"   : "Nama dan jabatan aktor utama / instansi / lokasi",
      "tone"           : "Positif" atau "Netral" atau "Negatif",
      "risiko_ais"     : "Risiko spesifik sesuai kategori: jika TATA KELOLA, sebutkan kelemahan pengendalian dan potensi penyimpangan konkret; jika PELAKSANAAN, sebutkan risiko gagal capaian atau pemborosan; jika EKSTERNAL, fokus pada dampak fiskal/sosial yang perlu diantisipasi",
      "area_perhatian" : "Aspek konkret yang perlu mendapat perhatian pengawasan BPKP — sebutkan objek, mekanisme, atau instansi yang perlu direviu atau diaudit"
    }

    Aturan ketat:
    - Tone HANYA salah satu dari: Positif, Netral, Negatif
    - risiko_ais dan area_perhatian harus SPESIFIK — hindari framing generik seperti "perlu transparansi" atau "perlu akuntabilitas"
    - Jika isu eksternal, tetap sebutkan siapa yang bertanggung jawab merespons di sisi pemerintah Indonesia
    - JANGAN mengarang kepanjangan singkatan — gunakan singkatan apa adanya sesuai Topik crawl yang diberikan
    - Jika konten terbatas, tetap isi semua field berdasarkan judul dan konteks topik crawl — JANGAN kosongkan field manapun
    - Bahasa Indonesia formal
    - Output HANYA JSON murni, tidak ada teks sebelum atau sesudah, tidak ada markdown fence"""

    # ── Analisis per artikel ───────────────────────────────────────────────
    def analisis_artikel(client: Groq, artikel: dict) -> dict:
        konten = str(artikel.get("snippet", "") or "").strip()
        konten_info = (
            f"Konten  : {konten}" if konten
            else "Konten  : [tidak tersedia — analisis berdasarkan judul dan topik crawl]"
        )
        prompt = (
            f"Topik crawl: {artikel.get('label_isu', '-')}\n"
            f"Judul   : {artikel['judul']}\n"
            f"Sumber  : {artikel['sumber']}\n"
            f"Tanggal : {artikel['tanggal']}\n"
            f"{konten_info}\n\nHasilkan JSON analisis."
        )
        try:
            resp = client.chat.completions.create(
                model="llama-3.3-70b-versatile",
                messages=[
                    {"role": "system", "content": PROMPT_SISTEM},
                    {"role": "user",   "content": prompt},
                ],
                temperature=0.3,
                max_tokens=800,
            )
            teks = resp.choices[0].message.content.strip()
            # Buang markdown fence
            teks = re.sub(r"^```json\s*|^```\s*|\s*```$", "", teks).strip()
            # Ambil hanya bagian JSON (dari { pertama sampai } terakhir)
            m = re.search(r"\{.*\}", teks, flags=re.DOTALL)
            if m:
                teks = m.group(0)
            hasil = json.loads(teks)
            if hasil.get("tone") not in ("Positif", "Netral", "Negatif"):
                hasil["tone"] = "Netral"
            return hasil
        except Exception as e:
            # Simpan pesan error untuk diagnosa (muncul di expander hasil)
            return {
                "ringkasan_isu": str(artikel.get("judul", "-")),
                "isu_subisu": "-", "aktor_lokasi": "-",
                "tone": "Netral", "risiko_ais": "-", "area_perhatian": "-",
                "_error": str(e)[:200],
            }

    # ── Excel builder ──────────────────────────────────────────────────────
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
            f"Total: {len(data)} artikel  |  Pusat Strategi Kebijakan Pengawasan BPKP"
        )
        style(c, C_SUB, sz=9)
        ws.row_dimensions[2].height = 16
        ws.row_dimensions[3].height = 5

        HEADERS = [
            "No", "Tanggal", "Sumber", "Link/Bukti", "Judul/Post",
            "Ringkasan Isu", "Isu/Subisu", "Aktor/Lokasi",
            "Tone Berita", "Risiko/Implikasi AIS", "Area Perhatian",
        ]
        for col, h in enumerate(HEADERS, 1):
            c = ws.cell(row=4, column=col, value=h)
            style(c, C_NAVY, bold=True, sz=10, center=True, fc=C_WHITE)
        ws.row_dimensions[4].height = 34

        for i, d in enumerate(data):
            r  = 5 + i
            bg = C_ODD if i % 2 == 0 else C_EVEN
            baris = [
                i + 1, d.get("tanggal","-"), d.get("sumber","-"),
                d.get("link","-"), d.get("judul","-"),
                d.get("ringkasan_isu","-"), d.get("isu_subisu","-"),
                d.get("aktor_lokasi","-"), d.get("tone","Netral"),
                d.get("risiko_ais","-"), d.get("area_perhatian","-"),
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

    # ── Sidebar konten crawl ───────────────────────────────────────────────
    with st.sidebar:
        st.markdown("### 📰 Media Crawl AIS")
        st.caption("Pusat Strategi Kebijakan Pengawasan BPKP")
        st.divider()

        groq_key_default = ""
        if hasattr(st, "secrets"):
            groq_key_default = st.secrets.get("GROQ_API_KEY", "")
        if groq_key_default:
            st.success("✅ API Key terbaca dari Secrets")
            groq_key = groq_key_default
        else:
            groq_key = st.text_input(
                "Groq API Key", type="password",
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
        run_btn = st.button("🔍 Mulai Crawl", use_container_width=True)

    # ── Main area crawl ────────────────────────────────────────────────────
    st.markdown("""
    <div class="main-header">
      <h1>📰 Analisis Isu Strategis Pengawasan</h1>
      <p>Media Crawl otomatis · Query Expansion · Analisis Groq · Output Excel · Pusat Strategi Kebijakan Pengawasan BPKP</p>
    </div>
    """, unsafe_allow_html=True)

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

        keywords_input = [k.strip() for k in re.split(r"[\n,]+", keywords_raw) if k.strip()]
        client = Groq(api_key=groq_key)

        st.subheader("⏳ Proses Crawl & Analisis")
        with st.spinner("Memperluas keyword dengan Groq..."):
            all_queries = []
            for kw in keywords_input:
                expanded = ekspansi_keyword(client, kw)
                all_queries.extend(expanded)

        query_lines = "\n".join(f"🔍 {q}" for q in all_queries)
        st.markdown(f'<div class="query-box"><b>Query yang akan dicari ({len(all_queries)} variasi):</b><br>{query_lines.replace(chr(10), "<br>")}</div>', unsafe_allow_html=True)

        prog_bar  = st.progress(0, text="Crawling Google News...")
        status_tx = st.empty()
        status_tx.info(f"Crawling {len(all_queries)} query ke Google News...")
        artikel_raw = crawl_google_news(all_queries, max_art)

        if not artikel_raw:
            st.warning("Tidak ada artikel ditemukan. Coba kata kunci yang berbeda atau lebih pendek.")
            st.stop()

        status_tx.success(f"✅ {len(artikel_raw)} artikel ditemukan. Memulai analisis Groq...")

        hasil_list = []
        for idx, art in enumerate(artikel_raw):
            pct = int((idx + 1) / len(artikel_raw) * 100)
            prog_bar.progress(pct, text=f"Menganalisis artikel {idx+1}/{len(artikel_raw)}...")
            time.sleep(0.15)
            art['label_isu'] = label_isu.strip()  # inject konteks topik ke tiap artikel
            analisis = analisis_artikel(client, art)
            hasil_list.append({**art, **analisis})

        prog_bar.progress(100, text="✅ Selesai!")
        status_tx.empty()

        st.session_state["hasil"]     = hasil_list
        st.session_state["label_isu"] = label_isu.strip()
        st.session_state["ais_ready"] = True  # flag untuk dashboard

        # Simpan error ke session agar tidak hilang saat re-run
        errors = [h.get("_error") for h in hasil_list if h.get("_error")]
        st.session_state["ais_errors"] = errors

        # Banner navigasi ke dashboard
        st.success("✅ Analisis selesai. Buka **📊 Dashboard AIS** di sidebar untuk melihat visualisasi lengkap.")

    # Tampilkan error diagnostik (persisten — tidak hilang saat re-run/download)
    if st.session_state.get("ais_errors"):
        errs = st.session_state["ais_errors"]
        total = len(st.session_state.get("hasil", []))
        with st.expander(f"⚠️ {len(errs)} dari {total} artikel gagal dianalisis — klik untuk detail", expanded=True):
            st.write("**Pesan error pertama:**")
            st.code(errs[0])
            low = errs[0].lower()
            if "rate" in low or "429" in low or "limit" in low:
                st.warning("🕐 Rate limit Groq. Tunggu 1-2 menit lalu coba lagi.")
            elif "auth" in low or "401" in low or "api" in low and "key" in low:
                st.warning("🔑 Masalah API Key Groq. Cek kembali key di Streamlit Secrets.")
            elif "json" in low or "expecting" in low:
                st.warning("📋 Groq mengembalikan respons non-JSON. Ini biasanya sementara — coba ulangi.")
            else:
                st.warning("Lihat pesan error di atas untuk diagnosa lebih lanjut.")

    if "hasil" in st.session_state:
        hasil_list = st.session_state["hasil"]
        label_isu  = st.session_state["label_isu"]

        tone_counts = {"Positif": 0, "Netral": 0, "Negatif": 0}
        for h in hasil_list:
            t = h.get("tone", "Netral")
            tone_counts[t] = tone_counts.get(t, 0) + 1

        c0, c1, c2, c3 = st.columns(4)
        with c0:
            st.markdown(f'<div class="stat-card"><div class="stat-number">{len(hasil_list)}</div><div class="stat-label">Total Artikel</div></div>', unsafe_allow_html=True)
        for col, tone, warna, emoji in [
            (c1, "Positif", "#065f46", "🟢"),
            (c2, "Netral",  "#92400e", "🟡"),
            (c3, "Negatif", "#991b1b", "🔴"),
        ]:
            with col:
                st.markdown(
                    f'<div class="stat-card">'
                    f'<div class="stat-number" style="color:{warna}">{tone_counts[tone]}</div>'
                    f'<div class="stat-label">{emoji} {tone}</div>'
                    f'</div>',
                    unsafe_allow_html=True,
                )

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("### ⬇️ Unduh Hasil")
        excel_buf = buat_excel(hasil_list, label_isu)
        nama_file = f"MediaCrawl_AIS_{label_isu.replace(' ','_')}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        st.download_button(
            "📥 Download Excel", data=excel_buf, file_name=nama_file,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("### 📋 Pratinjau Hasil Analisis")
        filter_tone = st.selectbox("Filter tone:", ["Semua", "Positif", "Netral", "Negatif"])
        tampil = hasil_list if filter_tone == "Semua" else [h for h in hasil_list if h.get("tone") == filter_tone]

        for h in tampil:
            tone = h.get("tone", "Netral")
            st.markdown(f"""
            <div class="artikel-card">
                <div class="artikel-judul">{h.get('judul','-')}</div>
                <div class="artikel-meta">📅 {h.get('tanggal','-')} &nbsp;·&nbsp; 📰 {h.get('sumber','-')} &nbsp;·&nbsp; {h.get('tier','')} &nbsp;·&nbsp; <span class="tone-{tone.lower()}">{tone}</span></div>
                <div class="artikel-ringkasan">{h.get('ringkasan_isu','-')}</div>
                <div class="artikel-risiko">⚠️ <b>Risiko:</b> {h.get('risiko_ais','-')}</div>
                <div class="artikel-tindak">🔍 <b>Area Perhatian:</b> {h.get('area_perhatian','-')}</div>
            </div>""", unsafe_allow_html=True)
            with st.expander("🔗 Lihat link artikel"):
                st.write(h.get("link", "-"))


# ══════════════════════════════════════════════════════════════════════════
# HALAMAN 2 — DASHBOARD AIS
# ══════════════════════════════════════════════════════════════════════════
elif page == "📊 Dashboard AIS":
    exec(open('dashboard_ais.py').read())
