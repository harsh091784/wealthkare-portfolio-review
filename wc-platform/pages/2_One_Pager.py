import streamlit as st
from pathlib import Path

ASSETS_DIR = Path(__file__).parent.parent / "assets"
LOGO_PATH = ASSETS_DIR / "logo_2.jpg"

TOOL_NAME = "One Pager"
TOOL_DESC = "Create a one-page client summary."

st.set_page_config(
    page_title=f"{TOOL_NAME} | Wealthkare",
    page_icon=str(LOGO_PATH) if LOGO_PATH.exists() else "📊",
    layout="centered",
)

# --- access gate -----------------------------------------------------------
# Every page is its own entry point in a Streamlit multipage app: a direct
# URL to /Portfolio_Review runs this file without app.py ever executing.
# The gate therefore has to be on each page, not only on the front one.
from pipeline.app_secrets import require_password  # noqa: E402

if not require_password():
    st.stop()
# ---------------------------------------------------------------------------

st.markdown(
    """
    <style>
        .wk-center { text-align: center; margin-top: 3rem; }
        .wk-title { color: #1C2B4B; font-weight: 700; font-size: 1.9rem; margin-bottom: 0.3rem; }
        .wk-desc { color: #444; font-size: 1rem; margin-bottom: 1.5rem; }
        .wk-badge {
            display: inline-block;
            background-color: #E5E0D5;
            color: #1C2B4B;
            font-weight: 600;
            font-size: 0.85rem;
            padding: 0.35rem 1rem;
            border-radius: 999px;
            margin-bottom: 1.2rem;
        }
        .wk-divider { border: none; border-top: 2px solid #B8860B; width: 120px; margin: 1.2rem auto; }
    </style>
    """,
    unsafe_allow_html=True,
)

col1, col2, col3 = st.columns([1, 1, 1])
with col2:
    if LOGO_PATH.exists():
        st.image(str(LOGO_PATH), use_container_width=True)

st.markdown(
    f"""
    <div class="wk-center">
        <div class="wk-badge">Coming Soon</div>
        <div class="wk-title">{TOOL_NAME}</div>
        <div class="wk-desc">{TOOL_DESC}</div>
        <hr class="wk-divider">
        <p style="color:#888; font-size:0.9rem;">This tool is not available yet.</p>
    </div>
    """,
    unsafe_allow_html=True,
)
