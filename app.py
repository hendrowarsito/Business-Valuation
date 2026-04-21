import streamlit as st
import os
from pathlib import Path

st.set_page_config(
    page_title="Lembar Kerja Penilaian Saham",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #1a5f4a 0%, #0d3d2e 100%);
        padding: 1.5rem 2rem;
        border-radius: 12px;
        margin-bottom: 1.5rem;
        color: white;
    }
    .main-header h1 { color: white; margin: 0; font-size: 1.6rem; }
    .main-header p { color: #9fe1cb; margin: 0.3rem 0 0 0; font-size: 0.85rem; }
    .step-card {
        background: white;
        border: 1px solid #e8f5f0;
        border-left: 4px solid #1D9E75;
        border-radius: 8px;
        padding: 1rem 1.25rem;
        margin-bottom: 0.75rem;
    }
    .step-card.inactive { border-left-color: #ddd; opacity: 0.6; }
    .status-ok { color: #1D9E75; font-weight: 600; }
    .status-warn { color: #BA7517; font-weight: 600; }
    .status-err { color: #E24B4A; font-weight: 600; }
    .metric-card {
        background: #f8fdf9;
        border: 1px solid #c8eadf;
        border-radius: 8px;
        padding: 0.75rem 1rem;
        text-align: center;
    }
    .metric-card .val { font-size: 1.8rem; font-weight: 700; color: #1a5f4a; }
    .metric-card .lbl { font-size: 0.75rem; color: #666; margin-top: 2px; }
    .tag-ok { background:#dcf5ec; color:#085041; padding:2px 8px; border-radius:12px; font-size:0.75rem; font-weight:600; }
    .tag-warn { background:#fef3d8; color:#633806; padding:2px 8px; border-radius:12px; font-size:0.75rem; font-weight:600; }
    .tag-formula { background:#dbeeff; color:#0c447c; padding:2px 8px; border-radius:12px; font-size:0.75rem; font-weight:600; }
    .tag-miss { background:#fde8e8; color:#791f1f; padding:2px 8px; border-radius:12px; font-size:0.75rem; font-weight:600; }
    .log-container { background:#0f1117; border-radius:8px; padding:1rem; font-family:monospace; font-size:0.78rem; max-height:250px; overflow-y:auto; }
    .log-ok { color:#4ade80; }
    .log-warn { color:#fbbf24; }
    .log-err { color:#f87171; }
    .log-info { color:#93c5fd; }
    .log-gray { color:#9ca3af; }
    .stButton > button { border-radius: 8px; }
    div[data-testid="stSidebarContent"] { background: #f8fdf9; }
</style>
""", unsafe_allow_html=True)

from utils.session import init_session
from pages import upload, analysis, review, download

init_session()

# ── Sidebar ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 📊 Penilaian Saham")
    st.markdown("---")
    
    steps = [
        ("1", "Upload Dokumen",    "upload"),
        ("2", "Analisis AI",       "analysis"),
        ("3", "Review Mapping",    "review"),
        ("4", "Download",          "download"),
    ]
    
    current = st.session_state.get("current_step", "upload")
    completed = st.session_state.get("completed_steps", [])
    
    for num, label, key in steps:
        icon = "✅" if key in completed else ("🔵" if key == current else "⚪")
        if key == current:
            st.markdown(f"**{icon} Step {num} — {label}**")
        else:
            st.markdown(f"{icon} Step {num} — {label}")
    
    st.markdown("---")
    st.markdown("**Pengaturan**")
    
    st.session_state.api_key = st.text_input(
        "Claude API Key", 
        value=st.session_state.get("api_key", os.getenv("ANTHROPIC_API_KEY", "")),
        type="password",
        help="Masukkan Claude API Key Anda dari console.anthropic.com"
    )
    
    st.session_state.currency = st.selectbox(
        "Mata Uang", ["IDR", "USD", "SGD"],
        index=["IDR","USD","SGD"].index(st.session_state.get("currency","IDR"))
    )
    
    st.session_state.unit = st.selectbox(
        "Satuan Angka", ["Juta", "Miliar", "Penuh"],
        index=["Juta","Miliar","Penuh"].index(st.session_state.get("unit","Juta"))
    )
    
    st.session_state.lang = st.selectbox(
        "Bahasa Akun", ["Indonesia", "Inggris", "Bilingual"],
        index=["Indonesia","Inggris","Bilingual"].index(st.session_state.get("lang","Indonesia"))
    )
    
    if st.button("🔄 Reset Semua", use_container_width=True):
        for key in list(st.session_state.keys()):
            if key not in ["api_key", "currency", "unit", "lang"]:
                del st.session_state[key]
        st.rerun()
    
    st.markdown("---")
    st.caption("v1.0 · SRR Appraisal · 2025")

# ── Main header ───────────────────────────────────────────────────────────────
st.markdown("""
<div class="main-header">
    <h1>📊 Lembar Kerja Penilaian Saham — Semi Otomatis</h1>
    <p>Berbasis Claude AI · openpyxl · Preservasi Format Template</p>
</div>
""", unsafe_allow_html=True)

# ── Route pages ───────────────────────────────────────────────────────────────
page = st.session_state.get("current_step", "upload")

if page == "upload":
    upload.render()
elif page == "analysis":
    analysis.render()
elif page == "review":
    review.render()
elif page == "download":
    download.render()
