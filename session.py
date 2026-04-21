import streamlit as st

def init_session():
    defaults = {
        "current_step": "upload",
        "completed_steps": [],
        "uploaded_files": {},
        "lk_data": None,
        "proj_data": None,
        "template_schema": None,
        "mapping_result": [],
        "log_messages": [],
        "api_key": "",
        "currency": "IDR",
        "unit": "Juta",
        "lang": "Indonesia",
        "output_path": None,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

def go_to(step: str):
    st.session_state.current_step = step

def mark_done(step: str):
    if step not in st.session_state.completed_steps:
        st.session_state.completed_steps.append(step)

def add_log(msg: str, level: str = "info"):
    st.session_state.log_messages.append({"msg": msg, "level": level})

def render_log():
    lines = st.session_state.get("log_messages", [])
    if not lines:
        return
    colors = {"ok": "#4ade80", "warn": "#fbbf24", "err": "#f87171", "info": "#93c5fd", "gray": "#9ca3af"}
    html = '<div class="log-container">'
    for entry in lines[-60:]:
        c = colors.get(entry["level"], "#9ca3af")
        html += f'<div style="color:{c}; margin-bottom:2px;">{entry["msg"]}</div>'
    html += "</div>"
    return html
