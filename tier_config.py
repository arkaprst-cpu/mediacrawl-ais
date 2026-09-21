"""Klasifikasi tier sumber media -- SATU sumber kebenaran dipakai bareng
oleh app.py (saat crawl) dan dashboard_ais.py (saat baca ulang Excel hasil
crawl). Sebelumnya daftar ini di-duplikasi manual di dua file dengan
komentar "harus sama persis" -- jebakan klasik: kalau satu diupdate dan
yang lain kelupaan, tier jadi beda antara crawl baru vs baca file lama
tanpa ada yang sadar. Sekarang cukup edit di sini, dua-duanya ikut.

Root cause 3 masalah nyata yang ditemukan pas diskusi (dicek langsung ke
data crawl asli, bukan cuma teori):
1. Daftar kata kuncinya belum lengkap -- CNBC Indonesia & MetroTVNews.com
   (dua-duanya media arus utama) sama sekali nggak ada di daftar lama.
2. Pencocokan substring gagal kalau nama sumbernya pakai SPASI, padahal
   kata kuncinya ditulis tanpa spasi -- "CNN Indonesia" (nama manusiawi
   dari suffix judul Google News) tidak mengandung "cnnindonesia" sebagai
   substring karena ada spasi, begitu juga "ANTARA News Jateng" vs
   "antaranews". Fix: normalisasi (buang semua spasi) sebelum dicocokkan.
3. Sumber non-media (universitas, kementerian, dll -- rilis resmi yang
   ikut ke-crawl Google News) dipaksa masuk Tier 1/2 padahal itu bukan
   kategori "media" sama sekali -- sekarang punya kategori sendiri.
"""

import re

# Kata kunci ditulis TANPA spasi/tanda baca (huruf kecil semua) --
# _normalize() di bawah juga membuang semua spasi/tanda baca dari nama
# sumber sebelum dicocokkan, jadi "CNN Indonesia" dan "cnnindonesia.com"
# dua-duanya jadi "cnnindonesia" dan cocok sama-sama.
TIER1_KEYWORDS = {
    "kompas", "tempo", "detik", "cnnindonesia", "republika", "antaranews",
    "mediaindonesia", "bisnis", "kontan", "tribunnews", "tribun", "liputan6",
    "okezone", "sindonews", "jpnn", "suara", "kumparan", "rmol", "inews",
    "katadata", "validnews", "thejakartapost", "jawapos",
    # Ditambah -- ditemukan langsung kelewat di data crawl asli (CNBC
    # Indonesia & MetroTVNews.com muncul beberapa kali, dua-duanya media
    # arus utama nasional, tapi jatuh ke Tier 2 karena absen dari daftar).
    "cnbcindonesia", "metrotvnews", "metrotv",
    # Ditambah -- outlet nasional lain yang jelas arus utama, belum ada
    # datanya di sampel yang dicek tapi kemungkinan besar bakal muncul di
    # crawl-crawl berikutnya.
    "merdeka", "viva", "beritasatu", "idntimes", "tirto", "gatra", "rri",
    "tvonenews", "antarafoto",
}

# Heuristik nama institusi/pemerintah -- rilis resmi dari lembaga ini yang
# ikut ke-crawl Google News BUKAN "media", jadi dipisah dari Tier 1/2
# (masuk Tier 1/2 sama-sama nggak pas -- lembaga negara bukan "media arus
# utama" ATAUPUN "sumber alternatif" dalam pengertian jurnalistik).
INSTITUSI_KEYWORDS = {
    "kementerian", "kemenko", "kemendagri", "kemenkeu", "kemenkes",
    "kemendikbud", "kemenag", "kemenhub", "kemenperin", "kemendag",
    "kemenparekraf", "kemenpora", "kemensos", "kemenkumham", "kemenlu",
    "kemenaker", "kementan", "kemenpupr", "kemenpanrb", "pendayagunaan",
    "reformasi birokrasi", "sekretariat negara", "sekretariat kabinet",
    "universitas", "institut teknologi", "politeknik", "perguruan tinggi",
    "pemerintah provinsi", "pemerintah kabupaten", "pemerintah kota",
    "pemprov", "pemkab", "pemkot", "dinas ", "badan pusat statistik", "bps",
    "dpr ri", "dprd", "mahkamah", "kejaksaan", "kepolisian", "polri", "polda",
    "ombudsman", "komisi pemberantasan korupsi", "bpkp", "bpk ri",
}


def pisah_judul_sumber(judul: str):
    """Pecah judul Google News "Judul Berita - NamaSumber" jadi
    (judul_tanpa_sumber, nama_sumber). Kalau pemisahnya tidak ketemu,
    return (judul_apa_adanya, "").

    Ditaruh di sini (bukan di app.py) karena dipakai bareng: app.py buat
    pill sumber & kartu Trending Topics, dashboard_ais.py buat membuang
    nama media sebelum tokenisasi kata kunci. Sebelumnya pola regexnya
    ditulis ulang manual di dua file -- persis jebakan yang dulu bikin
    classify_sumber() dipindah ke sini.

    Dulu nama sumber di-capture pakai [^-]+$ (tidak boleh mengandung
    tanda hubung) supaya judul yang SENDIRINYA punya " - " di tengah
    tidak ikut kepotong. Efek sampingnya: nama sumber yang memang
    bertanda hubung (mis. "portal-komando.com") gagal kedeteksi total --
    judulnya tampil mentah dengan ekor " - portal-komando.com" dan pill
    sumbernya hilang, dan di dashboard kata "portal"/"komando"/"com"
    ikut kehitung sebagai kata kunci isu. Sekarang dipotong di pemisah
    TERAKHIR: tujuan awalnya tetap kejaga (judul ber-" - " di tengah
    tidak kepotong, karena yang dipakai pemisah paling kanan), tapi nama
    sumber bertanda hubung ikut kebaca."""
    teks = str(judul).strip()
    # " - " dan " – " sama-sama 3 karakter (spasi + 1 tanda + spasi),
    # jadi offset +3 di bawah berlaku untuk dua-duanya.
    posisi = max(teks.rfind(" - "), teks.rfind(" – "))
    if posisi == -1:
        return teks, ""
    sumber = teks[posisi + 3:].strip()
    # Google News kadang menempelkan "[...]" di belakang nama sumber.
    sumber = re.sub(r"\s*\[.*?\]\s*$", "", sumber).strip()
    if not sumber:
        # Judulnya cuma kebetulan berakhir dengan " - " -- jangan dipotong.
        return teks, ""
    return teks[:posisi].strip(), sumber


def _normalize(nama: str) -> str:
    """Lowercase + buang semua spasi & tanda baca -- supaya "CNN Indonesia"
    (nama manusiawi dari suffix judul Google News) dan "cnnindonesia.com"
    (domain) sama-sama jadi bentuk yang bisa dicocokkan ke kata kunci."""
    return re.sub(r"[^a-z0-9]", "", str(nama).lower())


def classify_sumber(nama: str) -> str:
    """Klasifikasi satu nama sumber jadi "Tier 1" / "Tier 2" / "Institusi/
    Resmi". Institusi dicek DULUAN (pakai nama asli yang masih ada spasi,
    supaya "dinas " dengan spasi sengaja hanya cocok sebagai kata utuh,
    bukan substring nyasar) sebelum tier media dicek pakai versi yang
    sudah dinormalisasi."""
    nama_asli = str(nama).lower()
    if any(k in nama_asli for k in INSTITUSI_KEYWORDS):
        return "Institusi/Resmi"
    norm = _normalize(nama)
    if any(k in norm for k in TIER1_KEYWORDS):
        return "Tier 1"
    return "Tier 2"
