import streamlit as st
import pandas as pd
from utils.session import go_to, mark_done
from utils.writer import validate_balance_sheet


def render():
    st.markdown("### Step 3 — Review & Konfirmasi Mapping")

    mapping = st.session_state.get("mapping_result", [])
    lk_data = st.session_state.get("lk_data", {})
    proj_data = st.session_state.get("proj_data", {})

    if not mapping:
        st.error("Belum ada data mapping. Kembali ke Step 2.")
        if st.button("← Analisis"): go_to("analysis"); st.rerun()
        return

    # ── Info emiten ───────────────────────────────────────────────────────────
    emiten = lk_data.get("nama_emiten", "?")
    kode = lk_data.get("kode_saham", "?")
    tahun = lk_data.get("tahun", [])

    st.markdown(f"""
<div style="background:#f8fdf9;border:1px solid #c8eadf;border-radius:8px;
            padding:0.75rem 1.25rem;margin-bottom:1rem;display:flex;gap:2rem;align-items:center;">
    <div><span style="font-size:0.75rem;color:#666">Emiten</span><br>
         <strong>{emiten}</strong></div>
    <div><span style="font-size:0.75rem;color:#666">Kode</span><br>
         <strong>{kode}</strong></div>
    <div><span style="font-size:0.75rem;color:#666">Tahun Historis</span><br>
         <strong>{', '.join(str(y) for y in tahun)}</strong></div>
    <div><span style="font-size:0.75rem;color:#666">Satuan</span><br>
         <strong>{lk_data.get('satuan','Juta Rupiah')}</strong></div>
</div>
""", unsafe_allow_html=True)

    # ── Ringkasan stats ───────────────────────────────────────────────────────
    ok_items    = [m for m in mapping if m.get("tipe")=="input" and m.get("confidence",0)>=80]
    warn_items  = [m for m in mapping if m.get("tipe")=="input" and m.get("confidence",0)<80]
    formula_items = [m for m in mapping if m.get("tipe")=="formula"]
    miss_items  = [m for m in mapping if m.get("tipe")=="tidak_ditemukan"]

    c1,c2,c3,c4 = st.columns(4)
    with c1:
        st.markdown(f'<div class="metric-card"><div class="val">{len(mapping)}</div>'
                    f'<div class="lbl">Total Akun</div></div>', unsafe_allow_html=True)
    with c2:
        st.markdown(f'<div class="metric-card"><div class="val" style="color:#1D9E75">{len(ok_items)}</div>'
                    f'<div class="lbl">Berhasil Mapped</div></div>', unsafe_allow_html=True)
    with c3:
        st.markdown(f'<div class="metric-card"><div class="val" style="color:#BA7517">{len(warn_items)}</div>'
                    f'<div class="lbl">Perlu Review</div></div>', unsafe_allow_html=True)
    with c4:
        st.markdown(f'<div class="metric-card"><div class="val" style="color:#185FA5">{len(formula_items)}</div>'
                    f'<div class="lbl">Formula (Dijaga)</div></div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Validasi ──────────────────────────────────────────────────────────────
    validations = validate_balance_sheet(mapping)
    if validations:
        with st.expander("🔍 Hasil Validasi Otomatis", expanded=True):
            for v in validations:
                if v["status"] == "ok":
                    st.success(v["msg"])
                elif v["status"] == "warn":
                    st.warning(v["msg"])
                else:
                    st.error(v["msg"])

    # ── Tab filter ────────────────────────────────────────────────────────────
    tab_all, tab_ok, tab_warn, tab_formula, tab_miss = st.tabs([
        f"📋 Semua ({len(mapping)})",
        f"✅ Mapped ({len(ok_items)})",
        f"⚠️ Review ({len(warn_items)})",
        f"∑ Formula ({len(formula_items)})",
        f"❌ Tidak Ditemukan ({len(miss_items)})"
    ])

    def render_table(items, editable=False):
        if not items:
            st.markdown('<div class="empty-state" style="padding:1.5rem;text-align:center;'
                        'color:#999">Tidak ada data</div>', unsafe_allow_html=True)
            return

        rows = []
        for m in items:
            conf = m.get("confidence", 0)
            tipe = m.get("tipe", "input")

            if tipe == "formula":
                badge = "∑ Formula"
                badge_style = "tag-formula"
            elif conf >= 80:
                badge = "✅ Mapped"
                badge_style = "tag-ok"
            elif tipe == "tidak_ditemukan":
                badge = "❌ Tidak ada"
                badge_style = "tag-miss"
            else:
                badge = "⚠️ Review"
                badge_style = "tag-warn"

            rows.append({
                "Sheet": m.get("sheet", ""),
                "Sel": m.get("sel", ""),
                "Akun Sumber": m.get("nama_akun_sumber", ""),
                "Akun Template": m.get("nama_akun_template", ""),
                "Tahun": m.get("tahun", ""),
                "Nilai": m.get("nilai") if m.get("nilai") is not None else "—",
                "Conf%": conf if tipe != "formula" else "—",
                "Catatan": m.get("catatan", ""),
            })

        df = pd.DataFrame(rows)

        if editable and not df.empty:
            st.caption("💡 Anda dapat mengedit kolom **Nilai** secara manual sebelum generate.")
            edited = st.data_editor(
                df, use_container_width=True, hide_index=True,
                column_config={
                    "Sel":   st.column_config.TextColumn("Sel", width=70, disabled=True),
                    "Sheet": st.column_config.TextColumn("Sheet", width=80, disabled=True),
                    "Akun Sumber": st.column_config.TextColumn("Akun Sumber", width=200, disabled=True),
                    "Akun Template": st.column_config.TextColumn("Akun Template", width=200, disabled=True),
                    "Tahun": st.column_config.NumberColumn("Tahun", width=70, disabled=True),
                    "Nilai": st.column_config.TextColumn("Nilai", width=130),
                    "Conf%": st.column_config.NumberColumn("Conf %", width=70, disabled=True),
                    "Catatan": st.column_config.TextColumn("Catatan", width=200, disabled=True),
                }
            )
            # Apply edits back to session mapping
            for i, row in edited.iterrows():
                orig_idx = mapping.index(items[i]) if items[i] in mapping else -1
                if orig_idx >= 0 and str(row["Nilai"]) != str(items[i].get("nilai","—")):
                    try:
                        mapping[orig_idx]["nilai"] = float(str(row["Nilai"]).replace(",",""))
                    except (ValueError, TypeError):
                        mapping[orig_idx]["nilai"] = row["Nilai"]
            st.session_state.mapping_result = mapping
        else:
            st.dataframe(df, use_container_width=True, hide_index=True)

    with tab_all:
        render_table(mapping, editable=False)

    with tab_ok:
        render_table(ok_items, editable=True)

    with tab_warn:
        st.warning(f"⚠️  **{len(warn_items)} akun** dengan confidence < 80% memerlukan konfirmasi manual. "
                   "Edit nilai di kolom 'Nilai' jika diperlukan.")
        render_table(warn_items, editable=True)

    with tab_formula:
        st.info("∑ Sel-sel ini mengandung **formula Excel** dan **tidak akan diubah** saat generate.")
        render_table(formula_items, editable=False)

    with tab_miss:
        if miss_items:
            st.error(f"❌ **{len(miss_items)} akun** tidak ditemukan di template. "
                     "Akun-akun ini tidak akan diinput ke worksheet.")
        render_table(miss_items, editable=False)

    # ── Proyeksi ringkasan ────────────────────────────────────────────────────
    with st.expander("📈 Ringkasan Proyeksi & Valuasi"):
        proj = proj_data.get("proyeksi", {})
        val = proj_data.get("valuasi", {})
        tahun_proj = proj_data.get("tahun_proyeksi", [])

        if proj and tahun_proj:
            keys_to_show = ["pendapatan", "ebitda", "laba_bersih", "free_cash_flow"]
            rows = []
            for k in keys_to_show:
                if k in proj and proj[k]:
                    row = {"Keterangan": k.replace("_"," ").title()}
                    for i, y in enumerate(tahun_proj):
                        v = proj[k][i] if i < len(proj[k]) else None
                        row[str(y)] = f"{v:,.0f}" if v else "—"
                    rows.append(row)
            if rows:
                st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

        if val.get("nilai_wajar_per_saham"):
            cols = st.columns(3)
            with cols[0]:
                st.metric("Nilai Wajar/Saham", f"Rp {val['nilai_wajar_per_saham']:,.0f}")
            with cols[1]:
                harga = val.get("harga_pasar", 0)
                st.metric("Harga Pasar", f"Rp {harga:,.0f}" if harga else "—")
            with cols[2]:
                upside = val.get("upside_downside_pct")
                if upside is not None:
                    st.metric("Upside/Downside", f"{upside:+.1f}%",
                              delta=f"{upside:+.1f}%",
                              delta_color="normal")

    # ── Action buttons ────────────────────────────────────────────────────────
    st.markdown("---")
    col1, col2, col3 = st.columns([1, 1, 2])
    with col1:
        if st.button("← Ulangi Analisis"):
            for k in ["analysis_complete","analysis_done_steps"]:
                st.session_state.pop(k, None)
            go_to("analysis"); st.rerun()
    with col2:
        if st.button("🔄 Refresh Mapping"):
            st.rerun()
    with col3:
        if st.button("⚙️  Generate Lembar Kerja →", type="primary", use_container_width=True):
            mark_done("review")
            go_to("download")
            st.rerun()
