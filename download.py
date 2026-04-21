import streamlit as st
import io
from datetime import datetime
from utils.session import go_to, mark_done
from utils.writer import write_to_template, generate_review_log, validate_balance_sheet


def render():
    st.markdown("### Step 4 — Generate & Download Lembar Kerja")

    mapping = st.session_state.get("mapping_result", [])
    lk_data = st.session_state.get("lk_data", {})
    proj_data = st.session_state.get("proj_data", {})
    tpl_bytes = st.session_state.get("uploaded_files", {}).get("template")
    unit = st.session_state.get("unit", "Juta")

    if not mapping or not tpl_bytes:
        st.error("Data tidak lengkap. Kembali ke Step 3.")
        if st.button("← Review Mapping"):
            go_to("review"); st.rerun()
        return

    # ── Auto-generate jika belum ──────────────────────────────────────────────
    if not st.session_state.get("output_bytes"):
        with st.spinner("Mengisi template... mohon tunggu"):
            log_msgs = []
            def log(msg, level="info"):
                log_msgs.append({"msg": msg, "level": level})

            try:
                output_bytes = write_to_template(tpl_bytes, mapping, unit=unit, log_fn=log)
                st.session_state.output_bytes = output_bytes
                st.session_state.generate_log = log_msgs
                mark_done("download")
            except Exception as e:
                st.error(f"❌ Gagal generate: {e}")
                return

    output_bytes = st.session_state.output_bytes
    emiten = lk_data.get("nama_emiten", "Emiten")
    kode = lk_data.get("kode_saham", "SAHAM")
    ts = datetime.now().strftime("%Y%m%d_%H%M")
    filename = f"LK_Penilaian_{kode}_{ts}.xlsx"

    # ── Success banner ────────────────────────────────────────────────────────
    st.markdown(f"""
<div style="background:linear-gradient(135deg,#1a5f4a,#0d3d2e);border-radius:12px;
            padding:1.5rem 2rem;color:white;margin-bottom:1.5rem;">
    <h3 style="color:white;margin:0">✅ Lembar Kerja Berhasil Dibuat</h3>
    <p style="color:#9fe1cb;margin:0.4rem 0 0">
        {emiten} ({kode}) · Format & font template terjaga · Semua formula Excel aktif
    </p>
</div>
""", unsafe_allow_html=True)

    # ── Stats ─────────────────────────────────────────────────────────────────
    filled = sum(1 for m in mapping if m.get("tipe")=="input" and m.get("nilai") is not None)
    formula = sum(1 for m in mapping if m.get("tipe")=="formula")
    warn = sum(1 for m in mapping if m.get("tipe")=="input" and m.get("confidence",0)<80)
    miss = sum(1 for m in mapping if m.get("tipe")=="tidak_ditemukan")
    avg_conf = round(
        sum(m["confidence"] for m in mapping if m.get("confidence",0)>0) /
        max(1, sum(1 for m in mapping if m.get("confidence",0)>0))
    )

    c1,c2,c3,c4,c5 = st.columns(5)
    for col, label, value, color in [
        (c1, "Sel Terisi",        filled,   "#1D9E75"),
        (c2, "Formula Dijaga",    formula,  "#185FA5"),
        (c3, "Perlu Review",      warn,     "#BA7517"),
        (c4, "Tidak Ditemukan",   miss,     "#E24B4A"),
        (c5, "Avg. Confidence",   f"{avg_conf}%", "#1a5f4a"),
    ]:
        with col:
            st.markdown(f"""
<div class="metric-card">
    <div class="val" style="color:{color}">{value}</div>
    <div class="lbl">{label}</div>
</div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Download buttons ──────────────────────────────────────────────────────
    col_dl1, col_dl2, col_dl3 = st.columns(3)

    with col_dl1:
        st.download_button(
            label="⬇  Download Lembar Kerja (.xlsx)",
            data=output_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )

    review_log_text = generate_review_log(mapping, lk_data, proj_data)
    with col_dl2:
        st.download_button(
            label="📄  Download Review Log (.txt)",
            data=review_log_text.encode("utf-8"),
            file_name=f"review_log_{kode}_{ts}.txt",
            mime="text/plain",
            use_container_width=True
        )

    with col_dl3:
        if st.button("🔄  Proses Dokumen Baru", use_container_width=True):
            for k in ["uploaded_files","lk_data","proj_data","template_schema",
                      "mapping_result","analysis_complete","analysis_done_steps",
                      "log_messages","output_bytes","generate_log","completed_steps"]:
                st.session_state.pop(k, None)
            st.session_state.current_step = "upload"
            st.rerun()

    # ── Validasi final ────────────────────────────────────────────────────────
    st.markdown("---")
    with st.expander("🔍 Validasi Final", expanded=True):
        validations = validate_balance_sheet(mapping)
        for v in validations:
            if v["status"] == "ok":
                st.success(v["msg"])
            elif v["status"] == "warn":
                st.warning(v["msg"])
            else:
                st.error(v["msg"])

    # ── Review log ────────────────────────────────────────────────────────────
    with st.expander("📋 Review Log Mapping"):
        st.code(review_log_text, language="text")

    # ── Generate log ─────────────────────────────────────────────────────────
    gen_log = st.session_state.get("generate_log", [])
    if gen_log:
        with st.expander("⚙️  Log Generate"):
            colors = {"ok":"#4ade80","warn":"#fbbf24","err":"#f87171","info":"#93c5fd","gray":"#9ca3af"}
            html = '<div class="log-container">'
            for entry in gen_log:
                c = colors.get(entry["level"], "#9ca3af")
                html += f'<div style="color:{c};margin-bottom:2px;">{entry["msg"]}</div>'
            html += "</div>"
            st.markdown(html, unsafe_allow_html=True)

    # ── Ringkasan cara penggunaan ─────────────────────────────────────────────
    with st.expander("💡 Cara Membuka & Menggunakan File"):
        st.markdown(f"""
1. **Download** file `{filename}` menggunakan tombol di atas
2. **Buka** dengan Microsoft Excel atau LibreOffice Calc
3. **Cek** sel-sel yang di-highlight kuning — itu akun dengan confidence rendah
4. **Verifikasi** sel formula: Total Aset, Laba Kotor, dll sudah otomatis terhitung
5. **Input manual** akun yang tidak ditemukan (lihat Review Log)
6. **Simpan** file setelah selesai review

**Catatan penting:**
- Font, style, border, dan format angka template **tidak diubah**
- Formula Excel (=SUM, =B5+B6, dll) **tetap aktif dan berfungsi**
- Tidak ada baris atau kolom baru yang ditambahkan
        """)
