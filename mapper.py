import json, re
import anthropic


def _clean_json(raw: str) -> str:
    raw = raw.strip()
    raw = re.sub(r"^```json\s*", "", raw)
    raw = re.sub(r"^```\s*", "", raw)
    raw = re.sub(r"```$", "", raw)
    return raw.strip()


def _call(client: anthropic.Anthropic, system: str, prompt: str,
          pdf_b64: str = None, max_tokens: int = 4096) -> str:
    content = []
    if pdf_b64:
        content.append({
            "type": "document",
            "source": {"type": "base64", "media_type": "application/pdf", "data": pdf_b64}
        })
    content.append({"type": "text", "text": prompt})

    msg = client.messages.create(
        model="claude-opus-4-5",
        max_tokens=max_tokens,
        system=system,
        messages=[{"role": "user", "content": content}]
    )
    return msg.content[0].text


# ── Step 1: Extract financial statements ─────────────────────────────────────

EXTRACT_SYSTEM = """Anda adalah asisten ekstraksi data keuangan. 
Selalu kembalikan JSON yang valid saja. Tidak ada teks lain, tidak ada markdown, tidak ada penjelasan.
Semua angka dalam format numerik (bukan string). Gunakan null jika data tidak tersedia."""

LK_PROMPT = """Ekstrak SEMUA pos dari laporan keuangan berikut. 
Kembalikan JSON dengan format PERSIS seperti ini:
{{
  "nama_emiten": "...",
  "kode_saham": "...",
  "periode": "...",
  "tahun": [2021, 2022, 2023],
  "satuan": "Juta Rupiah",
  "neraca": {{
    "kas_setara_kas": [null, null, null],
    "investasi_jangka_pendek": [null, null, null],
    "piutang_usaha": [null, null, null],
    "piutang_lain": [null, null, null],
    "persediaan": [null, null, null],
    "biaya_dibayar_dimuka": [null, null, null],
    "pajak_dibayar_dimuka": [null, null, null],
    "aset_lancar_lain": [null, null, null],
    "total_aset_lancar": [null, null, null],
    "investasi_jangka_panjang": [null, null, null],
    "aset_tetap_bruto": [null, null, null],
    "akumulasi_depresiasi": [null, null, null],
    "aset_tetap_bersih": [null, null, null],
    "aset_tidak_berwujud": [null, null, null],
    "goodwill": [null, null, null],
    "aset_pajak_tangguhan": [null, null, null],
    "aset_tidak_lancar_lain": [null, null, null],
    "total_aset_tidak_lancar": [null, null, null],
    "total_aset": [null, null, null],
    "utang_usaha": [null, null, null],
    "utang_bank_jangka_pendek": [null, null, null],
    "utang_pajak": [null, null, null],
    "beban_akrual": [null, null, null],
    "utang_lancar_lain": [null, null, null],
    "total_liabilitas_jangka_pendek": [null, null, null],
    "utang_bank_jangka_panjang": [null, null, null],
    "obligasi": [null, null, null],
    "liabilitas_pajak_tangguhan": [null, null, null],
    "liabilitas_tidak_lancar_lain": [null, null, null],
    "total_liabilitas_jangka_panjang": [null, null, null],
    "total_liabilitas": [null, null, null],
    "modal_saham": [null, null, null],
    "tambahan_modal_disetor": [null, null, null],
    "saldo_laba": [null, null, null],
    "ekuitas_lain": [null, null, null],
    "total_ekuitas": [null, null, null],
    "total_liabilitas_ekuitas": [null, null, null]
  }},
  "laba_rugi": {{
    "pendapatan": [null, null, null],
    "hpp": [null, null, null],
    "laba_kotor": [null, null, null],
    "beban_penjualan": [null, null, null],
    "beban_umum_administrasi": [null, null, null],
    "beban_operasi_lain": [null, null, null],
    "total_beban_operasi": [null, null, null],
    "laba_operasi": [null, null, null],
    "pendapatan_bunga": [null, null, null],
    "beban_bunga": [null, null, null],
    "laba_rugi_kurs": [null, null, null],
    "pendapatan_lain": [null, null, null],
    "ebt": [null, null, null],
    "pajak_penghasilan": [null, null, null],
    "laba_bersih": [null, null, null],
    "laba_per_saham_dasar": [null, null, null],
    "ebitda": [null, null, null]
  }},
  "arus_kas": {{
    "kas_dari_operasi": [null, null, null],
    "kas_dari_investasi": [null, null, null],
    "kas_dari_pendanaan": [null, null, null],
    "kas_awal": [null, null, null],
    "kas_akhir": [null, null, null],
    "capex": [null, null, null],
    "free_cash_flow": [null, null, null]
  }},
  "rasio": {{
    "current_ratio": [null, null, null],
    "debt_to_equity": [null, null, null],
    "gross_margin": [null, null, null],
    "net_margin": [null, null, null],
    "roe": [null, null, null],
    "roa": [null, null, null]
  }}
}}

Data laporan keuangan:
{lk_text}"""


PROJ_PROMPT = """Ekstrak proyeksi keuangan dari data berikut.
Kembalikan JSON format PERSIS:
{{
  "tahun_proyeksi": [2024, 2025, 2026],
  "asumsi": {{
    "growth_pendapatan": [null, null, null],
    "gross_margin": [null, null, null],
    "ebitda_margin": [null, null, null],
    "net_margin": [null, null, null],
    "capex_persen_pendapatan": [null, null, null],
    "wacc": null,
    "terminal_growth": null
  }},
  "proyeksi": {{
    "pendapatan": [null, null, null],
    "hpp": [null, null, null],
    "laba_kotor": [null, null, null],
    "ebitda": [null, null, null],
    "laba_operasi": [null, null, null],
    "laba_bersih": [null, null, null],
    "capex": [null, null, null],
    "free_cash_flow": [null, null, null],
    "total_aset": [null, null, null],
    "total_ekuitas": [null, null, null]
  }},
  "valuasi": {{
    "metode": "DCF",
    "ev": null,
    "net_debt": null,
    "equity_value": null,
    "jumlah_saham_beredar": null,
    "nilai_wajar_per_saham": null,
    "harga_pasar": null,
    "upside_downside_pct": null
  }}
}}

Data proyeksi:
{proj_text}"""


# ── Step 2: Map akun → template cells ────────────────────────────────────────

MAPPING_SYSTEM = """Anda adalah ahli pemetaan data keuangan ke template Excel.
Kembalikan JSON array yang valid saja. Tidak ada teks lain."""

MAPPING_PROMPT = """Berdasarkan skema template Excel dan data keuangan berikut,
buat daftar mapping setiap akun keuangan ke sel target di template.

ATURAN PENTING:
1. Jangan pernah mengubah sel yang sudah mengandung FORMULA — tandai sebagai tipe "formula"
2. Hanya isi sel INPUT (nilai hardcode) 
3. Jika akun tidak ditemukan di template, tandai tipe "tidak_ditemukan"
4. confidence: 0-100 (seberapa yakin mapping ini benar)
5. Untuk sel multi-tahun, daftarkan satu baris per sel (per tahun)

Kembalikan JSON array:
[
  {{
    "akun_kunci": "kas_setara_kas",
    "nama_akun_sumber": "Kas dan Setara Kas",
    "nama_akun_template": "...",
    "sheet": "Sheet1",
    "sel": "C8",
    "tahun": 2023,
    "nilai": 1920000,
    "tipe": "input",
    "confidence": 95,
    "catatan": ""
  }},
  {{
    "akun_kunci": "total_aset_lancar",
    "nama_akun_sumber": "Total Aset Lancar",
    "nama_akun_template": "Total Aset Lancar",
    "sheet": "Sheet1",
    "sel": "C12",
    "tahun": 2023,
    "nilai": null,
    "tipe": "formula",
    "confidence": 100,
    "catatan": "Sel ini sudah mengandung formula Excel, tidak diubah"
  }}
]

SCHEMA TEMPLATE:
{template_schema}

DATA KEUANGAN (ringkasan):
Emiten: {nama_emiten}
Tahun historis: {tahun_historis}
Tahun proyeksi: {tahun_proyeksi}

Neraca (nilai terakhir / tahun terbaru):
{neraca_summary}

Laba Rugi (nilai terakhir):
{lr_summary}

Proyeksi:
{proj_summary}"""


# ── Public API ────────────────────────────────────────────────────────────────

def extract_lk(api_key: str, lk_text: str = None, pdf_b64: str = None,
               log_fn=None) -> dict:
    client = anthropic.Anthropic(api_key=api_key)
    prompt = LK_PROMPT.format(lk_text=lk_text or "(lihat dokumen PDF terlampir)")
    if log_fn:
        log_fn("⏳ Mengirim laporan keuangan ke Claude API...", "info")
    raw = _call(client, EXTRACT_SYSTEM, prompt, pdf_b64=pdf_b64, max_tokens=4096)
    if log_fn:
        log_fn("✅ Respons LK diterima dari API", "ok")
    try:
        data = json.loads(_clean_json(raw))
        if log_fn:
            emiten = data.get("nama_emiten", "?")
            tahun = data.get("tahun", [])
            log_fn(f"✅ Emiten: {emiten} | Tahun: {tahun}", "ok")
        return data
    except json.JSONDecodeError as e:
        if log_fn:
            log_fn(f"⚠️ JSON parse error LK: {e} — menggunakan data demo", "warn")
        return _demo_lk()


def extract_proj(api_key: str, proj_text: str, log_fn=None) -> dict:
    client = anthropic.Anthropic(api_key=api_key)
    prompt = PROJ_PROMPT.format(proj_text=proj_text)
    if log_fn:
        log_fn("⏳ Mengirim proyeksi keuangan ke Claude API...", "info")
    raw = _call(client, EXTRACT_SYSTEM, prompt, max_tokens=2048)
    if log_fn:
        log_fn("✅ Respons proyeksi diterima", "ok")
    try:
        data = json.loads(_clean_json(raw))
        if log_fn:
            ty = data.get("tahun_proyeksi", [])
            log_fn(f"✅ Proyeksi tahun: {ty}", "ok")
        return data
    except json.JSONDecodeError as e:
        if log_fn:
            log_fn(f"⚠️ JSON parse error proyeksi: {e} — menggunakan demo", "warn")
        return _demo_proj()


def map_accounts(api_key: str, lk_data: dict, proj_data: dict,
                 template_schema: str, log_fn=None) -> list:
    client = anthropic.Anthropic(api_key=api_key)

    neraca = lk_data.get("neraca", {})
    lr = lk_data.get("laba_rugi", {})
    proj = proj_data.get("proyeksi", {})
    tahun = lk_data.get("tahun", [])
    tahun_proj = proj_data.get("tahun_proyeksi", [])

    def fmt_dict(d: dict, keys: list) -> str:
        lines = []
        for k in keys:
            v = d.get(k)
            if v and any(x is not None for x in (v if isinstance(v, list) else [v])):
                lines.append(f"  {k}: {v}")
        return "\n".join(lines) if lines else "  (tidak ada data)"

    neraca_keys = ["kas_setara_kas","piutang_usaha","persediaan","total_aset_lancar",
                   "aset_tetap_bersih","total_aset","utang_usaha","total_liabilitas",
                   "total_ekuitas","total_liabilitas_ekuitas"]
    lr_keys = ["pendapatan","hpp","laba_kotor","laba_operasi","ebitda","laba_bersih"]
    proj_keys = ["pendapatan","ebitda","laba_bersih","free_cash_flow"]

    prompt = MAPPING_PROMPT.format(
        template_schema=template_schema[:6000],
        nama_emiten=lk_data.get("nama_emiten", "?"),
        tahun_historis=tahun,
        tahun_proyeksi=tahun_proj,
        neraca_summary=fmt_dict(neraca, neraca_keys),
        lr_summary=fmt_dict(lr, lr_keys),
        proj_summary=fmt_dict(proj, proj_keys)
    )

    if log_fn:
        log_fn("⏳ Menjalankan account mapping dengan Claude...", "info")
    raw = _call(client, MAPPING_SYSTEM, prompt, max_tokens=4096)
    if log_fn:
        log_fn("✅ Mapping selesai", "ok")

    try:
        result = json.loads(_clean_json(raw))
        if not isinstance(result, list):
            raise ValueError("bukan list")
        ok = sum(1 for r in result if r.get("tipe") == "input" and r.get("confidence", 0) >= 80)
        warn = sum(1 for r in result if r.get("tipe") == "input" and r.get("confidence", 0) < 80)
        formula = sum(1 for r in result if r.get("tipe") == "formula")
        if log_fn:
            log_fn(f"✅ Mapped: {ok} akun  |  ⚠️ Review: {warn}  |  ∑ Formula: {formula}", "ok")
        return result
    except (json.JSONDecodeError, ValueError) as e:
        if log_fn:
            log_fn(f"⚠️ Mapping parse error: {e} — generating demo mapping", "warn")
        return _demo_mapping(lk_data, proj_data)


# ── Demo data (fallback) ──────────────────────────────────────────────────────

def _demo_lk() -> dict:
    return {
        "nama_emiten": "PT. CONTOH TBK (Demo Data)",
        "kode_saham": "DEMO",
        "periode": "31 Desember 2023",
        "tahun": [2021, 2022, 2023],
        "satuan": "Juta Rupiah",
        "neraca": {
            "kas_setara_kas": [1250000, 1580000, 1920000],
            "piutang_usaha": [890000, 1020000, 1150000],
            "persediaan": [650000, 720000, 810000],
            "aset_lancar_lain": [320000, 380000, 420000],
            "total_aset_lancar": [3110000, 3700000, 4300000],
            "aset_tetap_bersih": [4500000, 4800000, 5100000],
            "aset_tidak_berwujud": [250000, 220000, 190000],
            "total_aset": [8060000, 8920000, 9790000],
            "utang_usaha": [720000, 830000, 950000],
            "utang_bank_jangka_pendek": [500000, 450000, 400000],
            "total_liabilitas_jangka_pendek": [1400000, 1520000, 1650000],
            "utang_bank_jangka_panjang": [2000000, 1800000, 1600000],
            "total_liabilitas": [3400000, 3320000, 3250000],
            "modal_saham": [2000000, 2000000, 2000000],
            "saldo_laba": [2660000, 3600000, 4540000],
            "total_ekuitas": [4820000, 5240000, 5870000],
            "total_liabilitas_ekuitas": [8060000, 8920000, 9790000],
        },
        "laba_rugi": {
            "pendapatan": [8500000, 9800000, 11200000],
            "hpp": [5100000, 5880000, 6720000],
            "laba_kotor": [3400000, 3920000, 4480000],
            "beban_penjualan": [900000, 1025000, 1150000],
            "beban_umum_administrasi": [900000, 1025000, 1150000],
            "laba_operasi": [1600000, 1870000, 2180000],
            "beban_bunga": [180000, 160000, 140000],
            "ebt": [1420000, 1710000, 2040000],
            "pajak_penghasilan": [355000, 427500, 510000],
            "laba_bersih": [1065000, 1282500, 1530000],
            "ebitda": [1850000, 2150000, 2480000],
        },
        "arus_kas": {
            "kas_dari_operasi": [1200000, 1450000, 1680000],
            "kas_dari_investasi": [-850000, -720000, -680000],
            "kas_dari_pendanaan": [-250000, -300000, -350000],
            "kas_akhir": [1250000, 1580000, 1920000],
            "capex": [850000, 720000, 680000],
            "free_cash_flow": [350000, 730000, 1000000],
        },
        "rasio": {
            "gross_margin": [0.40, 0.40, 0.40],
            "net_margin": [0.125, 0.131, 0.137],
            "roe": [0.221, 0.245, 0.261],
        }
    }


def _demo_proj() -> dict:
    return {
        "tahun_proyeksi": [2024, 2025, 2026],
        "asumsi": {
            "growth_pendapatan": [0.143, 0.133, 0.117],
            "ebitda_margin": [0.22, 0.23, 0.24],
            "wacc": 0.11,
            "terminal_growth": 0.04
        },
        "proyeksi": {
            "pendapatan": [12800000, 14500000, 16200000],
            "laba_kotor": [5120000, 5945000, 6804000],
            "ebitda": [2816000, 3335000, 3888000],
            "laba_bersih": [1750000, 2050000, 2380000],
            "capex": [750000, 800000, 850000],
            "free_cash_flow": [2066000, 2535000, 3038000],
        },
        "valuasi": {
            "metode": "DCF",
            "ev": 25000000,
            "net_debt": 2080000,
            "equity_value": 22920000,
            "jumlah_saham_beredar": 5000000,
            "nilai_wajar_per_saham": 4584,
            "harga_pasar": 3800,
            "upside_downside_pct": 20.6
        }
    }


def _demo_mapping(lk_data: dict, proj_data: dict) -> list:
    neraca = lk_data.get("neraca", {})
    lr = lk_data.get("laba_rugi", {})
    tahun = lk_data.get("tahun", [2021, 2022, 2023])
    t = tahun[-1] if tahun else 2023

    def last(key, src):
        v = src.get(key)
        if isinstance(v, list):
            vals = [x for x in v if x is not None]
            return vals[-1] if vals else None
        return v

    items = [
        ("kas_setara_kas",        "Kas dan Setara Kas",         "C8",  last("kas_setara_kas", neraca),       95),
        ("piutang_usaha",         "Piutang Usaha Bersih",        "C9",  last("piutang_usaha", neraca),        92),
        ("persediaan",            "Persediaan",                  "C10", last("persediaan", neraca),           94),
        ("aset_lancar_lain",      "Aset Lancar Lainnya",         "C11", last("aset_lancar_lain", neraca),     87),
        ("total_aset_lancar",     "Total Aset Lancar",           "C12", None,                                 100),
        ("aset_tetap_bersih",     "Aset Tetap - Neto",           "C15", last("aset_tetap_bersih", neraca),    96),
        ("aset_tidak_berwujud",   "Aset Tidak Berwujud",         "C16", last("aset_tidak_berwujud", neraca),  88),
        ("total_aset",            "TOTAL ASET",                  "C18", None,                                 100),
        ("utang_usaha",           "Utang Usaha",                 "C21", last("utang_usaha", neraca),          93),
        ("utang_bank_jangka_panjang","Utang Bank Jangka Panjang","C24", last("utang_bank_jangka_panjang", neraca), 91),
        ("total_liabilitas",      "TOTAL LIABILITAS",            "C26", None,                                 100),
        ("total_ekuitas",         "TOTAL EKUITAS",               "C28", None,                                 100),
        ("pendapatan",            "Pendapatan Bersih",           "C32", last("pendapatan", lr),               97),
        ("hpp",                   "Harga Pokok Penjualan",       "C33", last("hpp", lr),                      95),
        ("laba_kotor",            "Laba Kotor",                  "C34", None,                                 100),
        ("laba_operasi",          "Laba Usaha",                  "C37", last("laba_operasi", lr),             94),
        ("ebitda",                "EBITDA",                      "C38", last("ebitda", lr),                   82),
        ("laba_bersih",           "Laba Bersih",                 "C40", last("laba_bersih", lr),              96),
    ]

    result = []
    for akun_kunci, nama_template, sel, nilai, conf in items:
        is_formula = conf == 100 and nilai is None
        result.append({
            "akun_kunci": akun_kunci,
            "nama_akun_sumber": nama_template,
            "nama_akun_template": nama_template,
            "sheet": "Sheet1",
            "sel": sel,
            "tahun": t,
            "nilai": nilai,
            "tipe": "formula" if is_formula else "input",
            "confidence": conf,
            "catatan": "Formula Excel dijaga, tidak diubah" if is_formula else (
                "Nama akun berbeda, perlu konfirmasi" if conf < 90 else "")
        })
    return result
