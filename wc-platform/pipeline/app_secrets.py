"""
pipeline/app_secrets.py

One place that knows where the API key lives.

Order is st.secrets first, environment second. Streamlit Cloud has no
.env - secrets are injected through st.secrets - while a developer
running locally has a .env and no secrets.toml. Checking secrets first
means the deployed app uses the managed secret even if a stale key is
sitting in the process environment.

st.secrets raises rather than returning None when no secrets file exists
at all, which is normal locally, so that case is caught and treated as
"not configured here" rather than as an error.
"""

from __future__ import annotations

import os
from typing import Optional

API_KEY_NAME = "ANTHROPIC_API_KEY"


def get_api_key() -> Optional[str]:
    try:
        import streamlit as st
        value = st.secrets.get(API_KEY_NAME)
        if value:
            return str(value).strip()
    except Exception:
        pass  # no secrets.toml, or not running under Streamlit - fall through

    value = os.environ.get(API_KEY_NAME)
    return value.strip() if value else None


def build_anthropic_client(api_key: Optional[str] = None):
    """Returns a configured client, or None when no key is available.

    None is a supported state, not a failure: summary_client falls back to
    its deterministic template, so the report still builds with no key and
    no network.
    """
    key = api_key or get_api_key()
    if not key:
        return None
    import anthropic
    from pipeline.summary_client import API_MAX_RETRIES, API_TIMEOUT_SECONDS
    return anthropic.Anthropic(
        api_key=key, timeout=API_TIMEOUT_SECONDS, max_retries=API_MAX_RETRIES
    )


# --------------------------------------------------------------------------
# Shared-password gate
# --------------------------------------------------------------------------
# Streamlit Community Cloud serves every app on a public URL. This one
# ingests real client PANs, holdings and valuations and writes reports
# naming real people, so "public URL, unlisted" is not an access control.
# A single shared password is not strong authentication - it is the
# difference between "anyone who finds the link" and "someone who was
# given the password", which is the gap that matters here.

APP_PASSWORD_NAME = "APP_PASSWORD"
_AUTH_FLAG = "_authenticated"


def get_app_password() -> Optional[str]:
    try:
        import streamlit as st
        value = st.secrets.get(APP_PASSWORD_NAME)
        if value:
            return str(value)
    except Exception:
        pass
    value = os.environ.get(APP_PASSWORD_NAME)
    return value if value else None


def require_password() -> bool:
    """Blocks the page until the shared password is entered.

    Returns True when access is granted. When it returns False the caller
    must stop rendering - the gate has already drawn the password prompt.

    Once matched, the result is held in st.session_state so a rerun (every
    widget interaction is one) does not re-prompt.

    An UNCONFIGURED password fails closed. Deploying without setting
    APP_PASSWORD would otherwise leave the app wide open on a public URL,
    and a gate that quietly disables itself when misconfigured is worse
    than no gate, because it looks like protection.
    """
    import streamlit as st

    if st.session_state.get(_AUTH_FLAG):
        return True

    expected = get_app_password()
    if not expected:
        st.error(
            "**This app is not configured for access.**\n\n"
            f"No `{APP_PASSWORD_NAME}` is set. On Streamlit Cloud add it under "
            "**Settings → Secrets**; locally put it in `.streamlit/secrets.toml` "
            "or the environment. Access is refused until it is set."
        )
        return False

    st.markdown("#### Wealthkare — Portfolio Review")
    st.caption("This tool handles client data. Enter the access password to continue.")
    entered = st.text_input("Access password", type="password",
                            key="_password_input", label_visibility="collapsed")

    if not entered:
        return False

    # compare_digest keeps the check constant-time.
    import hmac
    if hmac.compare_digest(entered, expected):
        st.session_state[_AUTH_FLAG] = True
        # Don't leave the password sitting in session state.
        st.session_state.pop("_password_input", None)
        return True

    st.error("Incorrect password.")
    return False
