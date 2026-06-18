"""
Media Crawl AIS — Pusat Strategi Kebijakan Pengawasan BPKP
Streamlit web app: input keyword → query expansion → crawl → analisis → download Excel
Provider: DeepSeek (deepseek-v4-flash)
"""

import streamlit as st
import feedparser, json, time, re, io, os
from datetime import datetime
from urllib.parse import quote_plus
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Provider imports (lazy, agar tidak crash jika salah satu tidak terinstall) ──
def get_deepseek_client(api_key: str):
    from openai import OpenAI
    return OpenAI(api_key=api_key, base_url="https://api.deepseek.com")

st.set_page_config(
    page_title="Media Crawl AIS — Pusat Strategi Kebijakan Pengawasan",
    page_icon="📰",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── NAVIGASI ───────────────────────────────────────────────────────────────
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
    .artikel-judul { font-size: 1rem; font-weight: 600; color: #1e293b; margin-bottom: 0.4rem; line-height: 1.4; }
    .artikel-meta  { font-size: 0.78rem; color: #64748b; margin-bottom: 0.8rem; font-family: 'IBM Plex Mono', monospace; }
    .artikel-ringkasan { font-size: 0.88rem; color: #374151; line-height: 1.6; margin-bottom: 0.6rem; }
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
    .tone-positif { background:#d1fae5;color:#065f46;padding:2px 10px;border-radius:12px;font-size:0.75rem;font-weight:600; }
    .tone-netral  { background:#fef3c7;color:#92400e;padding:2px 10px;border-radius:12px;font-size:0.75rem;font-weight:600; }
    .tone-negatif { background:#fee2e2;color:#991b1b;padding:2px 10px;border-radius:12px;font-size:0.75rem;font-weight:600; }
    .query-box {
        background: #eff6ff; border: 1px solid #bfdbfe; border-radius: 8px;
        padding: 0.8rem 1rem; margin-bottom: 1rem;
        font-size: 0.83rem; color: #1e40af; font-family: 'IBM Plex Mono', monospace;
    }
    .provider-badge {
        display:inline-block; padding:2px 10px; border-radius:12px;
        font-size:0.72rem; font-weight:600; margin-left:6px;
    }
    .badge-deepseek { background:#fdf4ff; color:#86198f; border:1px solid #f5d0fe; }
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
        m = re.search(r"\s[-\u2013]\s([^-\u2013]+)$", judul.strip())
        if m:
            sumber = m.group(1).strip()
            sumber = re.sub(r"\s*\[.*?\]\s*$", "", sumber).strip()
            return sumber if sumber else ""
        return ""

    # ── PROMPT SISTEM ──────────────────────────────────────────────────────
    PROMPT_SISTEM = """Kamu adalah analis isu strategis pengawasan pemerintahan Indonesia untuk BPKP Pusat Strategi Kebijakan Pengawasan.

LANGKAH WAJIB — lakukan secara berurutan sebelum mengisi JSON:

LANGKAH 1 — BACA JUDUL SECARA LITERAL
Identifikasi: siapa yang disebut, apa yang terjadi, ada angka/besaran apa, ada kata kunci negatif apa (dugaan, korupsi, gagal, mangkrak, turun, naik, dll). Jangan tambahkan asumsi yang tidak ada di judul.

LANGKAH 2 — TENTUKAN PEMICU UTAMA
Pilih SATU kategori yang paling tepat berdasarkan isi judul:
- EKSTERNAL: pemicunya adalah faktor di luar kendali pemerintah Indonesia (harga global, geopolitik, bencana alam, kebijakan negara lain)
- KEBIJAKAN: ada keputusan atau regulasi pemerintah pusat/daerah yang menjadi objek berita
- PELAKSANAAN: ada program/proyek/anggaran yang bermasalah dalam implementasinya (lambat, mangkrak, tidak tepat sasaran, pemborosan)
- TATA KELOLA: ada indikasi penyimpangan, fraud, konflik kepentingan, atau kelemahan pengendalian internal

LANGKAH 3 — TENTUKAN SIAPA YANG BERTANGGUNG JAWAB
Sebutkan institusi/jabatan spesifik yang punya kewenangan atas isu ini.

LANGKAH 4 — ISI JSON
Gunakan hasil langkah 1–3. Wajib spesifik — DILARANG menulis frasa berikut: "perlu transparansi", "perlu akuntabilitas", "tata kelola yang baik", "penguatan pengawasan internal" (terlalu generik).

PENTING — perbedaan risiko vs area_perhatian:
- "risiko" menjawab: APA yang bisa terjadi jika kondisi ini tidak diintervensi (kerugian, kegagalan, penyimpangan). Ini adalah PERNYATAAN RISIKO, bukan rencana kerja. JANGAN diawali label kategori (TATA KELOLA/PELAKSANAAN/KEBIJAKAN/EKSTERNAL) — tulis langsung isinya.
- "area_perhatian" menjawab: TITIK LEMAH atau CELAH KONKRET apa yang melatarbelakangi risiko tersebut — bukan jenis kegiatan pengawasan yang harus dilakukan. JANGAN menulis "audit terhadap...", "reviu terhadap...", "perlu dilakukan pemeriksaan...", atau kalimat lain yang berbentuk rekomendasi/rencana tindakan pengawasan. Tulis sebagai temuan/celah, bukan instruksi kerja. BPKP yang akan menentukan sendiri bentuk pengawasannya — tugasmu hanya menunjukkan DI MANA letak titik lemahnya.

Contoh SALAH (area_perhatian berbentuk kegiatan pengawasan):
"Audit kinerja dan keuangan terhadap pengelolaan dapur MBG oleh BGN"

Contoh BENAR (area_perhatian berbentuk titik lemah):
"Standar kebersihan dan kualitas bahan baku pada dapur penyedia MBG belum terverifikasi secara independen, sementara pengelolaan dapur melibatkan banyak penyedia pihak ketiga dengan pengawasan harian yang minim dari BGN"

Output harus berupa JSON murni tanpa teks apapun di luar kurung kurawal:
{
  "ringkasan_isu"  : "2-3 kalimat: apa yang terjadi (sebutkan nama program/institusi/angka jika ada di judul), apa pemicunya, mengapa relevan bagi pengawasan BPKP",
  "isu_subisu"     : "Nama isu utama / subisu spesifik (gunakan istilah dari judul, bukan abstraksi)",
  "aktor_lokasi"   : "Nama institusi atau jabatan yang disebut dalam judul / lokasi spesifik",
  "tone"           : "Positif" atau "Netral" atau "Negatif",
  "risiko"         : "Pernyataan risiko konkret — apa yang bisa terjadi/dirugikan, pada objek apa, tanpa label kategori di depannya",
  "area_perhatian" : "Titik lemah atau celah konkret yang melatarbelakangi risiko tersebut — kondisi/struktur/mekanisme yang rentan, BUKAN jenis kegiatan audit/reviu yang harus dilakukan"
}

Aturan tambahan:
- Tone HANYA: Positif, Netral, atau Negatif
- Gunakan nama program/instansi/angka yang ada di judul — jangan ganti dengan abstraksi
- Jika judul tidak memberi cukup informasi, tetap isi semua field berdasarkan konteks topik crawl
- Bahasa Indonesia formal
- Output HANYA JSON murni"""

    # ── PROMPT KLASTER ─────────────────────────────────────────────────────
    PROMPT_KLASTER = """Kamu adalah analis isu strategis pengawasan BPKP Pusat Strategi Kebijakan Pengawasan.

Kamu akan menerima daftar artikel (no, judul, ringkasan_isu, isu_subisu) hasil crawl SATU keyword/topik yang sama.
Meski semua artikel membahas topik yang sama, arah/akar persoalannya bisa berbeda-beda.

TUGAS: kelompokkan artikel-artikel ini ke dalam klaster isu utama berdasarkan KESAMAAN AKAR PERSOALAN DAN ARAH ISU — bukan sekadar kesamaan kata kunci permukaan.

ATURAN:
1. Jumlah klaster FLEKSIBEL sesuai data — boleh 2, boleh 6, sesuaikan dengan keragaman isu yang benar-benar ada. Jangan memaksakan jumlah tertentu.
2. Setiap artikel HARUS masuk tepat satu klaster (tidak ada yang terlewat, tidak ada duplikasi).
3. Jangan membuat klaster "Lain-lain" kecuali benar-benar tidak ada kesamaan sama sekali dengan klaster lain.
4. Setiap klaster wajib diberi:
   - "nama": nama klaster singkat (maks 8 kata), mencerminkan isu utama bukan sekadar topik umum
   - "kondisi_pemicu": 1-2 kalimat kondisi/pemicu konkret yang menyatukan artikel-artikel ini
   - "akar_persoalan": akar tata kelola/kebijakan/pelaksanaan yang melatarbelakangi (bukan rekomendasi tindakan)
   - "risiko_utama": risiko paling signifikan jika tidak diintervensi
   - "relevansi_pengawasan": mengapa klaster ini relevan/tidak terlalu prioritas bagi pengawasan BPKP
   - "anggota": array berisi nomor (No) artikel yang masuk klaster ini

5. Urutkan array klaster dari yang paling kritikal/prioritas bagi pengawasan BPKP ke yang paling rendah prioritas.
6. Balas HANYA dalam format JSON murni, TANPA teks lain, TANPA markdown code fence, TANPA penjelasan di luar JSON.

Format output:
{"klaster": [{"nama": "...", "kondisi_pemicu": "...", "akar_persoalan": "...", "risiko_utama": "...", "relevansi_pengawasan": "...", "anggota": [1,2,3]}]}
"""

    def klasterisasi_isu_deepseek(client, hasil_list: list) -> list:
        """Kirim seluruh ringkasan isu ke DeepSeek untuk dikelompokkan jadi
        beberapa klaster isu utama. Mengembalikan list klaster (bisa kosong
        jika gagal — pemanggil wajib menangani fallback)."""
        ringkasan_list = [
            {
                "no": i + 1,
                "judul": h.get("judul", "-"),
                "ringkasan_isu": h.get("ringkasan_isu", "-"),
                "isu_subisu": h.get("isu_subisu", "-"),
            }
            for i, h in enumerate(hasil_list)
        ]
        prompt = json.dumps(ringkasan_list, ensure_ascii=False)

        MAX_RETRY  = 3
        BASE_DELAY = 5

        for attempt in range(MAX_RETRY):
            try:
                resp = client.chat.completions.create(
                    model="deepseek-v4-flash",
                    messages=[
                        {"role": "system", "content": PROMPT_KLASTER},
                        {"role": "user",   "content": prompt},
                    ],
                    temperature=0.2,
                    max_tokens=3000,
                )
                teks = resp.choices[0].message.content.strip()
                teks = re.sub(r"^```json\s*|^```\s*|\s*```$", "", teks).strip()
                m = re.search(r"\{.*\}", teks, flags=re.DOTALL)
                if m:
                    teks = m.group(0)
                parsed = json.loads(teks)
                klaster_list = parsed.get("klaster", [])

                # Validasi: pastikan setiap anggota klaster adalah int yang valid
                total_artikel = len(hasil_list)
                anggota_terpakai = set()
                klaster_valid = []
                for kl in klaster_list:
                    anggota = [a for a in kl.get("anggota", []) if isinstance(a, int) and 1 <= a <= total_artikel]
                    if not anggota:
                        continue
                    anggota_terpakai.update(anggota)
                    klaster_valid.append({**kl, "anggota": anggota})

                # Artikel yang tidak masuk klaster manapun -> klaster "Isu Lainnya"
                sisa = [n for n in range(1, total_artikel + 1) if n not in anggota_terpakai]
                if sisa:
                    klaster_valid.append({
                        "nama": "Isu Lainnya",
                        "kondisi_pemicu": "Artikel dengan arah isu yang tidak terkelompok ke klaster utama.",
                        "akar_persoalan": "-",
                        "risiko_utama": "-",
                        "relevansi_pengawasan": "Perlu ditelaah manual — tidak teridentifikasi pola yang jelas.",
                        "anggota": sisa,
                    })

                return klaster_valid

            except Exception:
                if attempt < MAX_RETRY - 1:
                    time.sleep(BASE_DELAY)
                    continue
                return []  # fallback: dashboard/Excel akan tampil tanpa klaster

        return []

    # ── Query expansion ────────────────────────────────────────────────────
    def ekspansi_keyword_deepseek(client, keyword: str) -> list:
        prompt = f"""Kamu adalah asisten pencarian berita. Dari input keyword berikut, buat 4-5 variasi query pencarian berita yang lebih spesifik dan efektif untuk Google News.

Keyword input: "{keyword}"

Aturan:
- Variasikan dengan sinonim, singkatan, nama lokasi spesifik, atau aspek berbeda dari isu yang sama
- Gunakan bahasa Indonesia
- Kembalikan HANYA array JSON berisi string query, tanpa teks lain

Contoh output: ["query 1", "query 2", "query 3", "query 4"]"""
        try:
            resp = client.chat.completions.create(
                model="deepseek-v4-flash",
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

    # ── Filter video ───────────────────────────────────────────────────────
    DOMAIN_VIDEO = {"kompas.tv","metrotvnews.com","tvone.co.id","rctiplus.com","vidio.com","youtube.com","youtu.be"}
    JUDUL_VIDEO  = ["[full]","[live]","[video]","[breaking]","live streaming","siaran langsung","tonton video","nonton:","breaking news:","full video"]

    def is_video(judul: str, domain: str) -> bool:
        judul_lower = judul.lower()
        for dv in DOMAIN_VIDEO:
            if dv in domain.lower(): return True
        for marker in JUDUL_VIDEO:
            if marker in judul_lower: return True
        return False

    def bersihkan_snippet(snippet: str, judul: str) -> str:
        teks = snippet.replace("&nbsp;"," ").replace("&amp;","&")
        teks = re.sub(r"&[a-z]+;"," ",teks)
        teks = re.sub(r"<[^>]+>","",teks)
        teks = re.sub(r"\s+"," ",teks).strip()
        judul_bersih = re.sub(r"\s+"," ",judul).strip().lower()
        if teks.lower().startswith(judul_bersih[:40].lower()):
            return ""
        return teks

    # ── Crawl Google News RSS ──────────────────────────────────────────────
    def crawl_google_news(queries: list, max_articles: int) -> list:
        articles = []
        seen     = set()
        skipped  = 0
        headers  = {"User-Agent": "Mozilla/5.0 (compatible; AIS-Crawler/1.0)"}

        for q in queries:
            url = f"https://news.google.com/rss/search?q={quote_plus(q)}&hl=id&gl=ID&ceid=ID:id"
            try:
                feed = feedparser.parse(url, request_headers=headers)
                for entry in feed.entries:
                    link = entry.get("link","")
                    if link in seen: continue
                    seen.add(link)

                    judul  = entry.get("title","-")
                    domain = re.sub(r"https?://(www\.)?","",link).split("/")[0]
                    if is_video(judul, domain):
                        skipped += 1
                        continue

                    try:
                        tanggal = datetime(*entry.published_parsed[:3]).strftime("%d %b %Y")
                    except Exception:
                        pub = entry.get("published","")
                        tanggal = pub[:10] if pub else "-"

                    konten      = bersihkan_snippet(re.sub(r"<[^>]+>","",entry.get("summary","")), judul)
                    sumber_nama = extract_sumber_dari_judul(judul) or re.sub(r"https?://(www\.)?","",link).split("/")[0]
                    tier_asli   = tier_sumber(sumber_nama.lower())

                    articles.append({
                        "judul": judul, "link": link, "tanggal": tanggal,
                        "sumber": sumber_nama, "snippet": konten, "tier": tier_asli,
                    })
                    if len(articles) >= max_articles:
                        if skipped > 0:
                            st.caption(f"ℹ️ {skipped} artikel video dilewati.")
                        return articles
            except Exception as e:
                st.warning(f"Gagal crawl '{q}': {e}")

        if skipped > 0:
            st.caption(f"ℹ️ {skipped} artikel video dilewati.")
        return articles

    # ── Progres tersimpan (mitigasi sesi Streamlit putus di tengah jalan) ──
    # Streamlit Cloud bisa memutus session_state kalau koneksi websocket
    # sempat terputus (tab idle lama, jaringan goyah, dll) — ini bukan bug
    # di kode, tapi keterbatasan platform. File ini menyimpan progres
    # sebagian supaya tidak hilang total kalau itu terjadi.
    PROGRES_FILE = "/tmp/ais_progres.json"

    def simpan_progres(label_isu: str, hasil_list: list, total: int):
        try:
            with open(PROGRES_FILE, "w", encoding="utf-8") as f:
                json.dump({
                    "label_isu": label_isu,
                    "total": total,
                    "selesai": len(hasil_list),
                    "hasil": hasil_list,
                    "timestamp": datetime.now().isoformat(),
                }, f, ensure_ascii=False)
        except Exception:
            pass  # progres gagal tersimpan tidak boleh menghentikan analisis

    def baca_progres():
        try:
            if os.path.exists(PROGRES_FILE):
                with open(PROGRES_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
        except Exception:
            pass
        return None

    def hapus_progres():
        try:
            if os.path.exists(PROGRES_FILE):
                os.remove(PROGRES_FILE)
        except Exception:
            pass

    # ── Analisis: DeepSeek ─────────────────────────────────────────────────
    def analisis_deepseek(client, artikel: dict, rate_status=None) -> dict:
        konten = str(artikel.get("snippet","") or "").strip()
        konten_info = f"Konten  : {konten}" if konten else "Konten  : [tidak tersedia — analisis berdasarkan judul dan topik crawl]"
        prompt = (
            f"Topik crawl: {artikel.get('label_isu','-')}\n"
            f"Judul   : {artikel['judul']}\n"
            f"Sumber  : {artikel['sumber']}\n"
            f"Tanggal : {artikel['tanggal']}\n"
            f"{konten_info}\n\nHasilkan JSON analisis."
        )

        MAX_RETRY  = 5
        BASE_DELAY = 6

        for attempt in range(MAX_RETRY):
            try:
                resp = client.chat.completions.create(
                    model="deepseek-v4-flash",
                    messages=[
                        {"role": "system", "content": PROMPT_SISTEM},
                        {"role": "user",   "content": prompt},
                    ],
                    temperature=0.3,
                    max_tokens=800,
                )
                konten_resp = (resp.choices[0].message.content or "").strip()

                # DeepSeek diketahui kadang membalas string KOSONG saat
                # rate limit tersembunyi terpicu (bukan error 429 eksplisit).
                # Ini harus ditangani sama seperti rate limit: backoff & retry,
                # bukan langsung gagal sebagai "JSON tidak valid".
                if not konten_resp:
                    wait = BASE_DELAY * (2 ** attempt)
                    if rate_status:
                        rate_status.warning(f"⏳ DeepSeek membalas kosong (indikasi rate limit) — menunggu {wait}s (retry {attempt+1}/{MAX_RETRY})...")
                    time.sleep(wait)
                    continue

                return _parse_json(konten_resp)

            except Exception as e:
                err_str = str(e).lower()
                if "429" in err_str or "rate" in err_str:
                    wait = BASE_DELAY * (2 ** attempt)
                    if rate_status:
                        rate_status.warning(f"⏳ Rate limit DeepSeek — menunggu {wait}s (retry {attempt+1}/{MAX_RETRY})...")
                    time.sleep(wait)
                elif attempt < MAX_RETRY - 1:
                    time.sleep(BASE_DELAY)
                    continue
                else:
                    return _fallback_error(artikel, str(e)[:200])

        return _fallback_error(artikel, "DeepSeek terus membalas kosong — kemungkinan rate limit tersembunyi, semua retry habis")

    # ── Helpers parse & fallback ───────────────────────────────────────────
    def _parse_json(teks: str) -> dict:
        teks = teks.strip()
        teks = re.sub(r"^```json\s*|^```\s*|\s*```$", "", teks).strip()
        m = re.search(r"\{.*\}", teks, flags=re.DOTALL)
        if m:
            teks = m.group(0)
        hasil = json.loads(teks)
        if hasil.get("tone") not in ("Positif", "Netral", "Negatif"):
            hasil["tone"] = "Netral"
        return hasil

    def _fallback_error(artikel: dict, pesan: str) -> dict:
        return {
            "ringkasan_isu": artikel.get("judul","-"),
            "isu_subisu": "-", "aktor_lokasi": "-",
            "tone": "Netral", "risiko": "-", "area_perhatian": "-",
            "_error": pesan,
        }

    # ── Excel builder ──────────────────────────────────────────────────────
    def buat_excel(data: list, label_isu: str) -> bytes:
        wb = Workbook()
        ws = wb.active
        ws.title = "Identifikasi Isu"

        C_NAVY="1F3864"; C_WHITE="FFFFFF"; C_SUB="D9E1F2"; C_ODD="EEF2F7"; C_EVEN="FFFFFF"
        TONE_C={"Positif":"C6EFCE","Netral":"FFEB9C","Negatif":"FFC7CE"}
        HEADERS=["No","Klaster Isu","Tanggal","Sumber","Link/Bukti","Judul/Post","Ringkasan Isu","Isu/Subisu","Aktor/Lokasi","Tone Berita","Risiko","Area Perhatian"]
        NCOL=len(HEADERS)
        s=Side(style="thin",color="CCCCCC")
        BD=Border(left=s,right=s,top=s,bottom=s)

        def style(c,bg,bold=False,sz=9,center=False,fc=C_NAVY):
            c.font=Font(name="Arial",size=sz,bold=bold,color=fc)
            c.fill=PatternFill("solid",fgColor=bg)
            c.alignment=Alignment(horizontal="center" if center else "left",vertical="top",wrap_text=True)
            c.border=BD

        ws.merge_cells(f"A1:{get_column_letter(NCOL)}1")
        c=ws["A1"]; c.value="IDENTIFIKASI ISU HARIAN — ANALISIS ISU STRATEGIS PENGAWASAN"
        style(c,C_NAVY,bold=True,sz=13,center=True,fc=C_WHITE); ws.row_dimensions[1].height=28

        ws.merge_cells(f"A2:{get_column_letter(NCOL)}2")
        c=ws["A2"]
        c.value=f"Isu: {label_isu}  |  Generate: {datetime.now().strftime('%d %B %Y, %H:%M')}  |  Total: {len(data)} artikel  |  Pusat Strategi Kebijakan Pengawasan BPKP"
        style(c,C_SUB,sz=9); ws.row_dimensions[2].height=16; ws.row_dimensions[3].height=5

        for col,h in enumerate(HEADERS,1):
            c=ws.cell(row=4,column=col,value=h); style(c,C_NAVY,bold=True,sz=10,center=True,fc=C_WHITE)
        ws.row_dimensions[4].height=34

        for i,d in enumerate(data):
            r=5+i; bg=C_ODD if i%2==0 else C_EVEN
            baris=[i+1,d.get("klaster","-"),d.get("tanggal","-"),d.get("sumber","-"),d.get("link","-"),d.get("judul","-"),
                   d.get("ringkasan_isu","-"),d.get("isu_subisu","-"),d.get("aktor_lokasi","-"),
                   d.get("tone","Netral"),d.get("risiko","-"),d.get("area_perhatian","-")]
            for col,val in enumerate(baris,1):
                c=ws.cell(row=r,column=col,value=val); style(c,bg)
            tone_val=d.get("tone","Netral")
            ws.cell(row=r,column=10).fill=PatternFill("solid",fgColor=TONE_C.get(tone_val,"FFEB9C"))
            ws.row_dimensions[r].height=60

        for col,w in enumerate([5,28,12,18,35,40,45,25,25,12,45,40],1):
            ws.column_dimensions[get_column_letter(col)].width=w

        buf=io.BytesIO(); wb.save(buf); buf.seek(0)
        return buf.read()

    # ══════════════════════════════════════════════════════════════════════
    # SIDEBAR — konfigurasi provider & input
    # ══════════════════════════════════════════════════════════════════════
    with st.sidebar:
        st.markdown("### 📰 Media Crawl AIS")
        st.caption("Pusat Strategi Kebijakan Pengawasan BPKP")
        st.divider()

        # ── API Key DeepSeek ─────────────────────────────────────────────
        # Provider tunggal: DeepSeek. Gemini & Groq sudah dicoba dan
        # tidak dipakai lagi — Gemini terlalu lambat (delay 6.5s/artikel
        # untuk hindari rate limit 10 RPM), Groq sering kosong di volume
        # tinggi karena rate limit ketat. DeepSeek terbukti cepat, murah
        # (<$0.01 per 20 artikel), dan hasil analisisnya konsisten.
        deepseek_key_default = st.secrets.get("DEEPSEEK_API_KEY","") if hasattr(st,"secrets") else ""
        if deepseek_key_default:
            st.success("✅ DeepSeek API Key dari Secrets")
            active_key = deepseek_key_default
        else:
            active_key = st.text_input("DeepSeek API Key", type="password", placeholder="sk-...")

        st.divider()

        keywords_raw = st.text_area(
            "Kata Kunci Isu",
            placeholder="Contoh:\nMBG, makan bergizi gratis\nDanantara ekspor\nPertamax BBM",
            height=130,
        )
        label_isu = st.text_input("Label Isu (nama file Excel)", placeholder="Contoh: Pertamax BBM Juni 2026")
        max_art   = st.slider("Maks. Artikel", min_value=5, max_value=20, value=20, step=5)

        # Tampilkan status cooldown jika masih dalam jeda
        COOLDOWN_DETIK = 180
        waktu_terakhir = st.session_state.get("crawl_selesai_at")
        if waktu_terakhir:
            selisih = time.time() - waktu_terakhir
            if selisih < COOLDOWN_DETIK:
                sisa = int(COOLDOWN_DETIK - selisih)
                st.warning(f"⏳ Cooldown: {sisa}s tersisa sebelum crawl baru bisa dimulai.")

        st.divider()
        run_btn = st.button("🔍 Mulai Crawl", use_container_width=True)

    # ── Main area ──────────────────────────────────────────────────────────

    # Deteksi progres tersimpan dari sesi yang terputus (websocket Streamlit
    # putus, tab idle lama, dll). Hanya relevan kalau session_state saat ini
    # kosong — kalau "hasil" sudah ada, berarti sesi masih hidup normal.
    progres_tersimpan = baca_progres() if "hasil" not in st.session_state else None
    if progres_tersimpan:
        sel = progres_tersimpan.get("selesai", 0)
        tot = progres_tersimpan.get("total", 0)
        st.warning(f"⚠️ Ditemukan progres crawl yang belum selesai: **{progres_tersimpan.get('label_isu','-')}** — {sel}/{tot} artikel sudah dianalisis sebelum sesi terputus.")
        c_pulih, c_buang = st.columns(2)
        with c_pulih:
            if st.button("📥 Pulihkan hasil sebagian ini", use_container_width=True):
                st.session_state["hasil"]      = progres_tersimpan["hasil"]
                st.session_state["klaster"]    = []
                st.session_state["label_isu"]  = progres_tersimpan.get("label_isu", "Hasil Crawl")
                st.session_state["ais_ready"]  = True
                st.session_state["ais_errors"] = [h.get("_error") for h in progres_tersimpan["hasil"] if h.get("_error")]
                st.rerun()
        with c_buang:
            if st.button("🗑️ Buang progres ini", use_container_width=True):
                hapus_progres()
                st.rerun()
        st.divider()

    provider_badge = '<span class="provider-badge badge-deepseek">DeepSeek V4 Flash</span>'
    st.markdown(f"""
    <div class="main-header">
      <h1>📰 Analisis Isu Strategis Pengawasan {provider_badge}</h1>
      <p>Media Crawl · Query Expansion · Analisis AI</p>
    </div>
    """, unsafe_allow_html=True)

    if run_btn:
        if not active_key:
            st.error("Masukkan API Key terlebih dahulu.")
            st.stop()
        if not keywords_raw.strip():
            st.error("Masukkan minimal satu kata kunci.")
            st.stop()
        if not label_isu.strip():
            st.error("Isi Label Isu untuk nama file Excel.")
            st.stop()

        # ── Cooldown antar-crawl ──────────────────────────────────────
        # Crawl beruntun tanpa jeda (mis. The Fed lalu langsung MBG)
        # membuat DeepSeek menganggap akun ini "burst" — limit dinamisnya
        # berdasarkan riwayat penggunaan jangka pendek, bukan cuma volume
        # 1 crawl saja. Cooldown ini memberi akun waktu "mendingin".
        COOLDOWN_DETIK = 180  # 3 menit
        waktu_terakhir = st.session_state.get("crawl_selesai_at")
        if waktu_terakhir:
            selisih = time.time() - waktu_terakhir
            if selisih < COOLDOWN_DETIK:
                sisa = int(COOLDOWN_DETIK - selisih)
                st.error(f"⏳ Tunggu {sisa} detik lagi sebelum crawl baru. Crawl beruntun tanpa jeda memicu rate limit DeepSeek lebih cepat — cooldown ini melindungi sesi crawl Anda sendiri.")
                st.stop()

        keywords_input = [k.strip() for k in re.split(r"[\n,]+", keywords_raw) if k.strip()]

        ai_client = get_deepseek_client(active_key)

        st.subheader("⏳ Proses Crawl & Analisis")

        with st.spinner("Memperluas keyword..."):
            all_queries = []
            for kw in keywords_input:
                expanded = ekspansi_keyword_deepseek(ai_client, kw)
                all_queries.extend(expanded)

        query_lines = "<br>".join(f"🔍 {q}" for q in all_queries)
        st.markdown(f'<div class="query-box"><b>Query ({len(all_queries)} variasi):</b><br>{query_lines}</div>', unsafe_allow_html=True)

        prog_bar  = st.progress(0, text="Crawling Google News...")
        status_tx = st.empty()
        status_tx.info(f"Crawling {len(all_queries)} query...")
        artikel_raw = crawl_google_news(all_queries, max_art)

        if not artikel_raw:
            st.warning("Tidak ada artikel ditemukan.")
            st.stop()

        status_tx.success(f"✅ {len(artikel_raw)} artikel ditemukan. Memulai analisis...")

        hasil_list  = []
        rate_status = st.empty()

        for idx, art in enumerate(artikel_raw):
            pct = int((idx + 1) / len(artikel_raw) * 100)
            prog_bar.progress(pct, text=f"Menganalisis artikel {idx+1}/{len(artikel_raw)}...")

            # DeepSeek menerapkan rate limit DINAMIS berbasis tekanan trafik
            # & pola burst (bukan RPM tetap yang dipublikasikan). Mengirim
            # 20+ request beruntun cepat lebih mudah memicu throttling
            # dibanding menyebar volume yang sama lebih lambat. Delay dasar
            # dinaikkan signifikan (3s) — ini PENCEGAHAN, bukan cuma reaksi
            # setelah gagal seperti backoff di analisis_deepseek().
            time.sleep(3.0)

            art["label_isu"] = label_isu.strip()
            analisis = analisis_deepseek(ai_client, art, rate_status)
            hasil_list.append({**art, **analisis})
            simpan_progres(label_isu.strip(), hasil_list, len(artikel_raw))

        rate_status.empty()
        prog_bar.progress(100, text="✅ Analisis selesai. Mengelompokkan isu...")

        with st.spinner("Mengelompokkan artikel menjadi klaster isu utama..."):
            klaster_list = klasterisasi_isu_deepseek(ai_client, hasil_list)

        # Tempel nama klaster ke setiap artikel (untuk Excel & tabel datar)
        klaster_per_no = {}
        for kl in klaster_list:
            for no in kl.get("anggota", []):
                klaster_per_no[no] = kl.get("nama", "-")
        for i, h in enumerate(hasil_list):
            h["klaster"] = klaster_per_no.get(i + 1, "-")

        prog_bar.progress(100, text="✅ Selesai!")
        status_tx.empty()

        st.session_state["hasil"]     = hasil_list
        st.session_state["klaster"]   = klaster_list
        st.session_state["label_isu"] = label_isu.strip()
        st.session_state["ais_ready"] = True
        st.session_state["ais_errors"] = [h.get("_error") for h in hasil_list if h.get("_error")]
        st.session_state["crawl_selesai_at"] = time.time()
        hapus_progres()  # crawl tuntas — progres sementara tidak diperlukan lagi

        if not klaster_list:
            st.warning("⚠️ Klasterisasi gagal — Excel & dashboard tetap tersedia tanpa pengelompokan isu.")
        st.success("✅ Analisis selesai. Buka **📊 Dashboard AIS** di sidebar untuk visualisasi lengkap.")

    # ── Error diagnostik ───────────────────────────────────────────────────
    if st.session_state.get("ais_errors"):
        errs  = st.session_state["ais_errors"]
        total = len(st.session_state.get("hasil", []))
        with st.expander(f"⚠️ {len(errs)} dari {total} artikel gagal dianalisis", expanded=True):
            st.code(errs[0])
            low = errs[0].lower()
            if "kosong" in low or "rate" in low or "429" in low or "quota" in low:
                st.warning("🕐 DeepSeek membalas kosong/rate limit di volume tinggi — retry otomatis dengan backoff sudah berjalan. Jika masih banyak yang gagal, tunggu 2–3 menit lalu ulangi, atau kecilkan jumlah Maks. Artikel per crawl.")
            elif "auth" in low or "401" in low:
                st.warning("🔑 Masalah API Key. Cek kembali key di Streamlit Secrets.")
            elif "json" in low or "expecting" in low:
                st.warning("📋 Respons non-JSON dari AI. Biasanya sementara — coba ulangi.")

    # ── Hasil & download ───────────────────────────────────────────────────
    if "hasil" in st.session_state:
        hasil_list = st.session_state["hasil"]
        label_isu  = st.session_state["label_isu"]

        tone_counts = {"Positif":0,"Netral":0,"Negatif":0}
        for h in hasil_list:
            t = h.get("tone","Netral")
            tone_counts[t] = tone_counts.get(t,0) + 1

        c0,c1,c2,c3 = st.columns(4)
        with c0:
            st.markdown(f'<div class="stat-card"><div class="stat-number">{len(hasil_list)}</div><div class="stat-label">Total Artikel</div></div>', unsafe_allow_html=True)
        for col,tone,warna,emoji in [(c1,"Positif","#065f46","🟢"),(c2,"Netral","#92400e","🟡"),(c3,"Negatif","#991b1b","🔴")]:
            with col:
                st.markdown(f'<div class="stat-card"><div class="stat-number" style="color:{warna}">{tone_counts[tone]}</div><div class="stat-label">{emoji} {tone}</div></div>', unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("### ⬇️ Unduh Hasil")
        excel_buf = buat_excel(hasil_list, label_isu)
        nama_file = f"MediaCrawl_AIS_{label_isu.replace(' ','_')}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        st.download_button("📥 Download Excel", data=excel_buf, file_name=nama_file,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("### 📋 Pratinjau Hasil Analisis")
        filter_tone = st.selectbox("Filter tone:", ["Semua","Positif","Netral","Negatif"])
        tampil = hasil_list if filter_tone=="Semua" else [h for h in hasil_list if h.get("tone")==filter_tone]

        for h in tampil:
            tone = h.get("tone","Netral")
            st.markdown(f"""
            <div class="artikel-card">
                <div class="artikel-judul">{h.get('judul','-')}</div>
                <div class="artikel-meta">📅 {h.get('tanggal','-')} &nbsp;·&nbsp; 📰 {h.get('sumber','-')} &nbsp;·&nbsp; {h.get('tier','')} &nbsp;·&nbsp; <span class="tone-{tone.lower()}">{tone}</span></div>
                <div class="artikel-ringkasan">{h.get('ringkasan_isu','-')}</div>
                <div class="artikel-risiko">⚠️ <b>Risiko:</b> {h.get('risiko','-')}</div>
                <div class="artikel-tindak">🔍 <b>Area Perhatian:</b> {h.get('area_perhatian','-')}</div>
            </div>""", unsafe_allow_html=True)
            with st.expander("🔗 Lihat link artikel"):
                st.write(h.get("link","-"))


# ══════════════════════════════════════════════════════════════════════════
# HALAMAN 2 — DASHBOARD AIS
# ══════════════════════════════════════════════════════════════════════════
elif page == "📊 Dashboard AIS":
    exec(open('dashboard_ais.py').read())
