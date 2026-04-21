import streamlit as st
from utils.session import go_to, mark_done, add_log, render_log
from utils import parser, mapper


def render():
    st.markdown("### Step 2 — Analisis AI")

    uploaded = st.session_state.get("uploaded_files", {})
    if not all(k in uploaded for k in ["lk", "proj", "template"]):
        st.error("Dokumen belum lengkap. Kembali ke Step 1.")
        if st.button("← Kembali ke Upload"):
            go_to("upload"); st.rerun()
        return

    api_key = st.session_state.get("api_key", "")
    if not api_key:
        st.warning("⚠️  Masukkan Claude API Key di sidebar untuk menjalankan analisis.")
        st.info("Tanpa API key, aplikasi akan menggunakan **data demo** untuk demonstrasi alur kerja.")

    # ── Status pipeline ───────────────────────────────────────────────────────
    step_labels = [
        ("parse_lk",    "1. Membaca & mengekstrak laporan keuangan"),
        ("parse_proj",  "2. Membaca & mengekstrak proyeksi keuangan"),
        ("parse_tpl",   "3. Menganalisis struktur template"),
        ("map_accounts","4. Mapping akun → sel template"),
    ]

    done_steps = st.session_state.get("analysis_done_steps", set())

    cols = st.columns(4)
    for i, (key, label) in enumerate(step_labels):
        with cols[i]:
            if key in done_steps:
                st.success(f"✅ {label.split('. ')[1][:25]}")
            elif len(done_steps) == i:
                st.info(f"⏳ {label.split('. ')[1][:25]}")
            else:
                st.markdown(f"<div style='color:#aaa;font-size:0.8rem'>⚪ {label.split('. ')[1][:25]}</div>",
                            unsafe_allow_html=True)

    # ── Log display ───────────────────────────────────────────────────────────
    log_placeholder = st.empty()
    log_html = render_log()
    if log_html:
        log_placeholder.markdown(log_html, unsafe_allow_html=True)

    # ── Already done ──────────────────────────────────────────────────────────
    if st.session_state.get("analysis_complete"):
        st.success("✅ Analisis selesai! Klik tombol di bawah untuk review mapping.")
        col1, col2 = st.columns([1, 4])
        with col1:
            if st.button("Lihat Hasil →", type="primary"):
                mark_done("analysis"); go_to("review"); st.rerun()
        with col2:
            if st.button("🔄 Ulangi Analisis"):
                for k in ["analysis_complete","analysis_done_steps","lk_data","proj_data",
                          "template_schema","mapping_result","log_messages"]:
                    st.session_state.pop(k, None)
                st.rerun()
        return

    # ── Run button ────────────────────────────────────────────────────────────
    if st.button("▶  Jalankan Analisis AI", type="primary"):
        _run_analysis(uploaded, api_key, log_placeholder)


def _run_analysis(uploaded, api_key, log_placeholder):
    done_steps = set()
    st.session_state.log_messages = []
    st.session_state.analysis_done_steps = done_steps

    def log(msg, level="info"):
        add_log(msg, level)
        log_placeholder.markdown(render_log(), unsafe_allow_html=True)

    # ── Step 1: Parse LK ─────────────────────────────────────────────────────
    log("━━━ Step 1: Ekstraksi Laporan Keuangan ━━━", "gray")
    lk_bytes = uploaded["lk"]
    is_pdf = st.session_state.get("lk_is_pdf", False)
    lk_data = None

    if api_key:
        try:
            if is_pdf:
                log("📄 Mode: PDF → Claude Vision", "info")
                pdf_b64 = parser.pdf_to_base64(lk_bytes)
                lk_data = mapper.extract_lk(api_key, pdf_b64=pdf_b64, log_fn=log)
            else:
                log("📊 Mode: XLSX → text parsing", "info")
                lk_text = parser.xlsx_to_text(lk_bytes)
                lk_data = mapper.extract_lk(api_key, lk_text=lk_text, log_fn=log)
        except Exception as e:
            log(f"❌ Error API LK: {e} — fallback ke demo", "err")
    
    if not lk_data:
        log("⚠️  Menggunakan data demo LK (tidak ada API key atau error)", "warn")
        lk_data = mapper._demo_lk()

    st.session_state.lk_data = lk_data
    done_steps.add("parse_lk")
    st.session_state.analysis_done_steps = done_steps
    st.rerun()

    # ── Step 2: Parse Proyeksi ────────────────────────────────────────────────
    log("━━━ Step 2: Ekstraksi Proyeksi Keuangan ━━━", "gray")
    proj_bytes = uploaded["proj"]
    proj_data = None

    if api_key:
        try:
            proj_text = parser.xlsx_to_text(proj_bytes)
            proj_data = mapper.extract_proj(api_key, proj_text=proj_text, log_fn=log)
        except Exception as e:
            log(f"❌ Error API Proyeksi: {e} — fallback ke demo", "err")

    if not proj_data:
        log("⚠️  Menggunakan data demo proyeksi", "warn")
        proj_data = mapper._demo_proj()

    st.session_state.proj_data = proj_data
    done_steps.add("parse_proj")
    st.session_state.analysis_done_steps = done_steps

    # ── Step 3: Parse Template ────────────────────────────────────────────────
    log("━━━ Step 3: Analisis Struktur Template ━━━", "gray")
    tpl_bytes = uploaded["template"]

    try:
        schema = parser.read_template_schema(tpl_bytes)
        formula_cells = parser.get_formula_cells(schema)
        label_map = parser.get_label_map(schema)
        st.session_state.template_schema = schema

        log(f"✅ Template dibaca: {len(schema['sheets'])} sheet", "ok")
        log(f"   Sel label ditemukan: {len(label_map)}", "info")
        log(f"   Sel formula (dilindungi): {len(formula_cells)}", "info")

        for sname in [s["name"] for s in schema["sheets"]]:
            log(f"   Sheet: {sname}", "gray")
    except Exception as e:
        log(f"❌ Error membaca template: {e}", "err")
        schema = {"sheets": [], "raw_text": ""}
        st.session_state.template_schema = schema

    done_steps.add("parse_tpl")
    st.session_state.analysis_done_steps = done_steps

    # ── Step 4: Account Mapping ───────────────────────────────────────────────
    log("━━━ Step 4: Account Mapping ━━━", "gray")
    mapping_result = []

    if api_key:
        try:
            mapping_result = mapper.map_accounts(
                api_key, lk_data, proj_data,
                schema.get("raw_text", ""),
                log_fn=log
            )
        except Exception as e:
            log(f"❌ Error mapping API: {e} — fallback ke demo mapping", "err")

    if not mapping_result:
        log("⚠️  Menggunakan demo mapping", "warn")
        mapping_result = mapper._demo_mapping(lk_data, proj_data)

    st.session_state.mapping_result = mapping_result
    done_steps.add("map_accounts")
    st.session_state.analysis_done_steps = done_steps

    ok = sum(1 for m in mapping_result if m.get("tipe")=="input" and m.get("confidence",0)>=80)
    warn = sum(1 for m in mapping_result if m.get("tipe")=="input" and m.get("confidence",0)<80)
    formula = sum(1 for m in mapping_result if m.get("tipe")=="formula")
    miss = sum(1 for m in mapping_result if m.get("tipe")=="tidak_ditemukan")

    log("━━━ Selesai ━━━", "gray")
    log(f"✅ Mapped: {ok}  |  ⚠️ Review: {warn}  |  ∑ Formula: {formula}  |  ❌ Tidak ditemukan: {miss}", "ok")

    st.session_state.analysis_complete = True
    st.rerun()
