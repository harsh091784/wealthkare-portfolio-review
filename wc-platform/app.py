import streamlit as st
from pathlib import Path

ASSETS_DIR = Path(__file__).parent / "assets"
LOGO_PATH = ASSETS_DIR / "logo_2.jpg"

st.set_page_config(
    page_title="Wealthkare | WC Securities",
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
        /* Tighten Streamlit's own top padding so the brand block sits higher */
        .block-container {
            padding-top: 2.2rem;
            padding-bottom: 3rem;
            max-width: 760px;
        }

        .wk-logo-wrap {
            max-width: 340px;
            margin: 0 auto 0.4rem auto;
        }

        .wk-title {
            color: #1C2B4B;
            font-weight: 800;
            font-size: 2.15rem;
            letter-spacing: -0.01em;
            margin-bottom: 0.15rem;
            text-align: center;
        }
        .wk-subtitle {
            color: #B8860B;
            font-weight: 600;
            font-size: 0.95rem;
            letter-spacing: 0.04em;
            text-transform: uppercase;
            margin-bottom: 1.6rem;
            text-align: center;
        }
        .wk-divider {
            border: none;
            height: 3px;
            width: 100%;
            margin: 0 0 1.8rem 0;
            background: linear-gradient(90deg, #B8860B 0%, #D9B04A 45%, rgba(184,134,11,0) 100%);
            border-radius: 3px;
        }
        .wk-intro {
            color: #555;
            font-size: 0.95rem;
            margin-bottom: 1.4rem;
        }

        .wk-card {
            position: relative;
            border: 1px solid #E7E2D6;
            border-left: 4px solid #1C2B4B;
            background-color: #FDFDFB;
            border-radius: 12px;
            padding: 1.1rem 1.3rem 1.1rem 1.5rem;
            margin-bottom: 1rem;
            box-shadow: 0 1px 3px rgba(28, 43, 75, 0.06);
            transition: transform 0.16s ease, box-shadow 0.16s ease, border-left-color 0.16s ease;
        }
        .wk-card:hover {
            transform: translateY(-3px);
            box-shadow: 0 10px 24px rgba(184, 134, 11, 0.18), 0 2px 6px rgba(28, 43, 75, 0.08);
            border-left-color: #B8860B;
        }
        .wk-card-title {
            color: #1C2B4B;
            font-weight: 700;
            font-size: 1.08rem;
            margin-bottom: 0.35rem;
            display: flex;
            align-items: center;
            gap: 0.6rem;
        }
        .wk-card-desc {
            color: #5A5A5A;
            font-size: 0.9rem;
            line-height: 1.4;
        }
        .wk-status {
            display: inline-block;
            font-size: 0.68rem;
            font-weight: 700;
            letter-spacing: 0.03em;
            text-transform: uppercase;
            padding: 0.2rem 0.6rem;
            border-radius: 999px;
            white-space: nowrap;
        }
        .wk-status-live {
            background-color: #B8860B;
            color: #FFFFFF;
        }
        .wk-status-soon {
            background-color: #F0EEE6;
            color: #8A7B4E;
            border: 1px solid #E0D6B8;
        }
        .wk-footer {
            text-align: center;
            color: #A0A0A0;
            font-size: 0.76rem;
            margin-top: 2.8rem;
            padding-top: 1.2rem;
            border-top: 1px solid #EDEDE8;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

col1, col2, col3 = st.columns([1, 2, 1])
with col2:
    if LOGO_PATH.exists():
        st.markdown("<div class='wk-logo-wrap'>", unsafe_allow_html=True)
        st.image(str(LOGO_PATH), use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

st.markdown("<div class='wk-title'>Wealthkare Internal Tools</div>", unsafe_allow_html=True)
st.markdown(
    "<div class='wk-subtitle'>WC Securities Pvt Ltd &nbsp;&middot;&nbsp; ARN: 3511</div>",
    unsafe_allow_html=True,
)
st.markdown("<div class='wk-divider'></div>", unsafe_allow_html=True)

st.markdown("<div class='wk-intro'>Select a tool from the sidebar, or a card below, to get started.</div>", unsafe_allow_html=True)

tools = [
    ("Portfolio Review Report Generator", "Generate client portfolio review reports.", "soon"),
    ("One Pager", "Create a one-page client summary.", "soon"),
    ("NSDL One Pager", "Create a one-page NSDL holdings summary.", "soon"),
    ("Brokerage Recon", "Reconcile brokerage statements.", "soon"),
    ("Market Update", "Generate market update briefs.", "soon"),
]

for name, desc, status in tools:
    badge_class = "wk-status-live" if status == "live" else "wk-status-soon"
    badge_text = "Live" if status == "live" else "Coming Soon"
    st.markdown(
        f"""
        <div class="wk-card">
            <div class="wk-card-title">{name}<span class="wk-status {badge_class}">{badge_text}</span></div>
            <div class="wk-card-desc">{desc}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

st.markdown(
    "<div class='wk-footer'>Wealthkare &mdash; Relationships Beyond Investments &nbsp;|&nbsp; Internal Use Only</div>",
    unsafe_allow_html=True,
)
