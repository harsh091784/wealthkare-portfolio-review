"""
Portfolio Review Report Generator - three screens, one wizard.

  1  Upload      parse the dashboard export, show what was found
  2  Review      pick a client, resolve warnings into decisions
  3  Approve     read the AI draft, approve it, generate the report

Deliberately ONE page with step routing rather than three files in
pages/. Streamlit's sidebar would otherwise let someone open "Approve"
with nothing parsed, and every screen depends on the one before it.

Nothing here computes a figure. The page parses, collects human
decisions, and calls the pipeline.
"""

from __future__ import annotations

import time
import traceback
from datetime import date
from pathlib import Path

import streamlit as st

from pipeline.app_secrets import build_anthropic_client, get_api_key
from pipeline.chart_gen import format_inr
from pipeline.dashboard_parser import ParseError, parse_dashboard_workbook
from pipeline.docx_builder import RMInfo
from pipeline.report_assembler import assemble_report_context
from pipeline.risk_profile import MARKET_CAP_RULES

ASSETS_DIR = Path(__file__).parent.parent / "assets"
LOGO_PATH = ASSETS_DIR / "logo_2.jpg"

st.set_page_config(page_title="Portfolio Review | Wealthkare",
                   page_icon=str(LOGO_PATH) if LOGO_PATH.exists() else "📊",
                   layout="wide")

# --- access gate -----------------------------------------------------------
# Every page is its own entry point in a Streamlit multipage app: a direct
# URL to /Portfolio_Review runs this file without app.py ever executing.
# The gate therefore has to be on each page, not only on the front one.
from pipeline.app_secrets import require_password  # noqa: E402

if not require_password():
    st.stop()
# ---------------------------------------------------------------------------

st.markdown("""
<style>
  .block-container { padding-top: 2rem; max-width: 1100px; }
  /* line-height and padding are load-bearing, not cosmetic: at 0.85rem
     uppercase the default line-height clips the glyph box, and
     letter-spacing adds a trailing space after the final character that
     the container was cutting off - the label rendered as "Step 1 of ".
     nowrap stops "of 3" wrapping onto a second, clipped line. */
  .wk-step { color:#8a8a8a; font-size:0.85rem; letter-spacing:0.08em; text-transform:uppercase;
             line-height:1.6; padding:0.15rem 0.08em 0.15rem 0; white-space:nowrap;
             overflow:visible; }
  .wk-h    { color:#1C2B4B; font-weight:700; font-size:1.6rem; margin:0.1rem 0 1rem 0; }
  .wk-draft{ background:#FFF8E6; border-left:4px solid #B8860B; padding:0.75rem 1rem;
             border-radius:4px; margin-bottom:0.5rem; }
  /* the .docx download is the secondary route - present, quieter */
  div[data-testid="stDownloadButton"] button[kind="secondary"] {
      background:transparent; color:#5a5a5a; border:1px solid #d5d5d5; font-weight:400;
  }
</style>
""", unsafe_allow_html=True)

MARKET_CAP_CHOICES = ["Unclassified"] + sorted({label for _kw, label in MARKET_CAP_RULES})

S = st.session_state
S.setdefault("step", 1)
S.setdefault("parse_result", None)
S.setdefault("parse_error", None)
S.setdefault("selected_client", None)
S.setdefault("warnings_ack", False)
S.setdefault("overrides", {})
S.setdefault("assembled", None)
S.setdefault("summary_draft", None)
S.setdefault("summary_source", None)
S.setdefault("deliverables", None)
S.setdefault("pdf_unavailable_reason", None)


def header(step: int, title: str) -> None:
    st.markdown(f'<div class="wk-step">Step {step} of 3</div>'
                f'<div class="wk-h">{title}</div>', unsafe_allow_html=True)


def reset_from(step: int) -> None:
    """Changing an earlier answer invalidates everything downstream. Kept
    explicit so a re-upload can never leave a previous client's approved
    summary attached to a new file."""
    if step <= 1:
        S.parse_result = S.parse_error = S.selected_client = None
    if step <= 2:
        S.warnings_ack = False
        S.overrides = {}
        S.assembled = None
    S.summary_draft = S.summary_source = S.deliverables = None


# ==========================================================================
# STEP 1 - Upload
# ==========================================================================

if S.step == 1:
    header(1, "Upload the dashboard export")
    st.caption("The .xlsx exported from the dashboard, after ACTION and "
               "SUGGESTED SCHEME have been filled in.")

    uploaded = st.file_uploader("Portfolio workbook", type=["xlsx"],
                                accept_multiple_files=False)

    if uploaded is not None:
        signature = (uploaded.name, uploaded.size)
        if S.get("upload_signature") != signature:
            S.upload_signature = signature
            reset_from(1)
            work = Path("_uploads"); work.mkdir(exist_ok=True)
            path = work / uploaded.name
            path.write_bytes(uploaded.getbuffer())
            try:
                S.parse_result = parse_dashboard_workbook(path)
            except ParseError as exc:
                S.parse_error = str(exc)
            except Exception as exc:  # unexpected - show it rather than a blank screen
                S.parse_error = f"{type(exc).__name__}: {exc}"
                S.parse_traceback = traceback.format_exc()

    if S.parse_error:
        st.error(f"**This file could not be parsed.**\n\n{S.parse_error}")
        st.caption("Nothing has been imported. Fix the file and upload it again.")
        if S.get("parse_traceback"):
            with st.expander("Technical detail"):
                st.code(S.parse_traceback)

    elif S.parse_result is not None:
        result = S.parse_result
        st.success(f"Parsed **{result.source_name}** — "
                   f"{len(result.clients)} client(s), "
                   f"{sum(len(c.holdings) for c in result.clients)} holdings.")
        st.caption(f"Sheets read: {', '.join(result.sheet_names)}")

        st.dataframe(
            [{
                "Client": c.name,
                "Holdings": len(c.holdings),
                "Current value": format_inr(c.computed_grand_total),
                "Actions": len(c.actions),
                "SIPs": len(c.sips),
                "Grand total check": {True: "reconciles", False: "MISMATCH",
                                      None: "no total row on sheet"}[c.grand_total_reconciles],
                "Warnings": len(c.warnings),
            } for c in result.clients],
            use_container_width=True, hide_index=True,
        )

        for warning in result.warnings:
            st.warning(warning.describe())

        if st.button("Continue →", type="primary"):
            S.step = 2
            st.rerun()


# ==========================================================================
# STEP 2 - Select client and review
# ==========================================================================

elif S.step == 2:
    header(2, "Select a client and review what was found")
    result = S.parse_result
    if result is None:
        st.warning("No file loaded.")
        st.stop()

    names = result.client_names()
    chosen = st.selectbox("Client", names,
                          index=names.index(S.selected_client) if S.selected_client in names else 0)
    if chosen != S.selected_client:
        S.selected_client = chosen
        reset_from(2)

    client = result.client(chosen)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Holdings", len(client.holdings))
    c2.metric("Purchase total", format_inr(client.total_purchase_value))
    c3.metric("Current value", format_inr(client.computed_grand_total))
    reconciles = client.grand_total_reconciles
    c4.metric("Grand total check",
              {True: "Reconciles", False: "Mismatch", None: "Not on sheet"}[reconciles])

    if reconciles is False:
        st.error(f"Holdings sum to {format_inr(client.computed_grand_total)} but the sheet's "
                 f"Grand Total says {format_inr(client.reported_grand_total)}.")
    elif reconciles is None:
        st.info("This sheet carries no Grand Total row, so the reconciliation check could not "
                "be run. That is not the same as it passing.")

    st.markdown("#### Actions found")
    if client.actions:
        st.dataframe(
            [{"Scheme": a.scheme, "Folio": a.folio or "—", "Action": a.action_raw,
              "Goes to": {"transaction": "Transaction Snapshot",
                          "things_to_do": "Things To Do",
                          "unrecognised": "NOT APPLIED — needs a decision"}[a.kind],
              "Suggested scheme": a.suggested_scheme or "—"} for a in client.actions],
            use_container_width=True, hide_index=True,
        )
    else:
        st.info("No actions on this client. The report will build with an empty Transaction "
                "Snapshot and no Mind Map section.")

    # ---- warnings become decisions --------------------------------------
    st.markdown("#### Warnings")
    blocking = list(client.warnings)

    # Assemble once so market-cap classification failures surface here too.
    preview = assemble_report_context(
        client, as_of=date.today(),
        rm=RMInfo(name="Relationship Manager", email="rm@wealthcareindia.com",
                  phone="+91-98100-00000"),
    )
    unclassified = preview.unclassified_schemes

    if not blocking and not unclassified:
        st.success("No warnings for this client.")
    else:
        for warning in blocking:
            st.warning(warning.describe())

        if unclassified:
            st.markdown("**Schemes the market-cap rules could not classify**")
            st.caption("Assign a category, or accept Unclassified. This choice affects the "
                       "Equity Sub-Allocation table only.")
            for scheme in unclassified:
                S.overrides[scheme] = st.selectbox(
                    scheme, MARKET_CAP_CHOICES,
                    index=MARKET_CAP_CHOICES.index(S.overrides.get(scheme, "Unclassified")),
                    key=f"override::{scheme}",
                )

        S.warnings_ack = st.checkbox(
            "I have reviewed the warnings above and want to continue",
            value=S.warnings_ack,
        )

    can_proceed = (not blocking and not unclassified) or S.warnings_ack

    left, right = st.columns([1, 4])
    if left.button("← Back"):
        S.step = 1
        st.rerun()
    if right.button("Proceed →", type="primary", disabled=not can_proceed):
        S.assembled = assemble_report_context(
            client, as_of=date.today(),
            rm=RMInfo(name="Relationship Manager", email="rm@wealthcareindia.com",
                      phone="+91-98100-00000"),
            market_cap_overrides={k: v for k, v in S.overrides.items() if v != "Unclassified"},
        )
        S.step = 3
        st.rerun()
    if not can_proceed:
        right.caption("Acknowledge the warnings to continue.")


# ==========================================================================
# STEP 3 - Summary approval and generate
# ==========================================================================

elif S.step == 3:
    header(3, "Review the summary, then generate")
    assembled = S.assembled
    if assembled is None:
        st.warning("No client selected.")
        st.stop()

    ctx = assembled.ctx
    st.caption(f"{ctx.client_name} — {len(ctx.holdings)} holdings, "
               f"{format_inr(ctx.portfolio_summary.current_value)}")

    if S.summary_draft is None:
        with st.spinner("Generating the client summary…"):
            from pipeline.summary_client import build_summary_input, generate_client_summary
            payload = build_summary_input(ctx)
            generated = generate_client_summary(payload, client=build_anthropic_client())
            S.summary_draft = generated.text
            S.summary_source = generated.source
            S.summary_failures = generated.failure_log

    if S.summary_source == "fallback":
        reason = "no API key is configured" if not get_api_key() else "the API call did not succeed"
        st.info(f"This is the **deterministic fallback** summary, not model-written text — "
                f"{reason}. It is built from the same computed figures and is safe to send, "
                f"but it reads plainly. Edit it as you would any draft.")
        if S.get("summary_failures"):
            with st.expander("Why the model was not used"):
                for entry in S.summary_failures:
                    st.write(f"- {entry}")

    st.markdown('<div class="wk-draft"><strong>AI-generated draft — requires review.</strong><br>'
                'Every figure has been checked against the computed data, but the wording has '
                'not been read by anyone yet. Edit it below before approving.</div>',
                unsafe_allow_html=True)

    S.summary_draft = st.text_area("Client Summary", value=S.summary_draft, height=280,
                                   label_visibility="collapsed")

    approved = st.checkbox("I have reviewed and approve this summary", value=False)

    # Say what this environment can actually produce, before the click
    # rather than after it. On Streamlit Community Cloud LibreOffice is
    # frequently absent, and finding that out from a red error box after
    # approving a summary is a worse experience than being told up front.
    from pipeline.pdf_converter import soffice_available
    if not soffice_available():
        st.warning(
            "**PDF rendering is unavailable in this environment** (LibreOffice is not "
            "installed). Generating will produce the editable **.docx only**, with "
            "placeholder dots on the contents page. Everything else in the report is "
            "complete."
        )

    st.divider()
    left, right = st.columns([1, 4])
    if left.button("← Back"):
        S.step = 2
        S.deliverables = None
        st.rerun()

    if right.button("Generate report", type="primary", disabled=not approved):
        from pipeline.mindmap import generate_mindmap
        from pipeline.pdf_converter import PdfUnavailable, build_report_deliverables
        from pipeline.summary_client import ClientSummary

        progress = st.progress(0.0, text="Preparing…")
        workdir = Path("_generated") / ctx.client_name.replace(" ", "_")
        workdir.mkdir(parents=True, exist_ok=True)
        try:
            if assembled.mindmap_recommendations:
                progress.progress(0.15, text="Drawing the mind map…")
                mindmap = workdir / "mindmap.png"
                generate_mindmap(assembled.mindmap_recommendations,
                                 client_name=ctx.client_name, output_path=mindmap)
                ctx.mindmap_path = mindmap

            ctx.client_summary = ClientSummary(text=S.summary_draft, approved=True,
                                               source=S.summary_source or "fallback")
            ctx.allow_missing_summary = False   # a summary is mandatory from here on

            progress.progress(0.35, text="Building the document…")
            started = time.perf_counter()
            try:
                pdf_path, docx_path, _pages = build_report_deliverables(ctx, workdir)
            except PdfUnavailable as unavailable:
                # No LibreOffice in this environment. The .docx is built
                # and complete - hand it over rather than failing the run
                # because the secondary format could not be produced.
                pdf_path, docx_path = None, unavailable.docx_path
                S.pdf_unavailable_reason = str(unavailable)
            else:
                S.pdf_unavailable_reason = None
            progress.progress(1.0, text=f"Done in {time.perf_counter() - started:.1f}s")
            S.deliverables = (pdf_path, docx_path)
        except Exception as exc:
            progress.empty()
            st.error(f"**Report generation failed.**\n\n{type(exc).__name__}: {exc}")
            with st.expander("Technical detail"):
                st.code(traceback.format_exc())

    if not approved:
        right.caption("Approve the summary to enable generation.")

    if S.deliverables:
        pdf_path, docx_path = S.deliverables
        if pdf_path is None:
            st.warning(
                "**Report generated as .docx only — no PDF.**\n\n"
                f"{S.get('pdf_unavailable_reason', '')}\n\n"
                "The contents page shows placeholder dots instead of page numbers. "
                "Open the .docx in Word and use *References → Update Table*, then export "
                "to PDF from there."
            )
        else:
            st.success("Report generated.")
        d1, d2 = st.columns([1, 1])
        with d1:
            if pdf_path is not None:
                st.download_button("⬇  Download PDF", data=Path(pdf_path).read_bytes(),
                                   file_name=Path(pdf_path).name, mime="application/pdf",
                                   type="primary", use_container_width=True)
        with d2:
            st.download_button(
                "⬇  Download .docx" if pdf_path is None else "Download .docx (editable)",
                data=Path(docx_path).read_bytes(), file_name=Path(docx_path).name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                # When it is the only deliverable it is the primary one.
                type="secondary" if pdf_path is not None else "primary",
                use_container_width=True)
