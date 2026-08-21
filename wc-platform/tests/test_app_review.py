"""
tests/test_app_review.py

AppTest coverage for screens 2 and 3 of the Portfolio Review page.

The approval gate is the only thing standing between a model-written
draft and a client, so it is tested here rather than clicked through: a
gate that is verified by hand is a gate that is verified until the day
somebody is in a hurry.

These drive the real page script through Streamlit's AppTest harness -
the same widgets, the same session state, the same reruns - with session
state seeded to land directly on the screen under test.
"""

from __future__ import annotations

from datetime import date
from pathlib import Path

import pytest
from streamlit.testing.v1 import AppTest

from pipeline.dashboard_parser import parse_dashboard_workbook
from pipeline.report_assembler import assemble_report_context
from pipeline.docx_builder import RMInfo

PAGE = Path(__file__).parent.parent / "pages" / "1_Portfolio_Review.py"
REAL_DATA_DIR = Path(__file__).parent / "data"

# Files are located by shape, never by filename - the real exports are
# named after the clients they belong to. The working sample here is the
# two-client, 28-holding book: small enough to render quickly, and it
# carries the paired switches worth exercising.
SAMPLE_SHAPE = (2, 28)
RM = RMInfo(name="Relationship Manager", email="rm@wealthcareindia.com", phone="+91-98100-00000")


def _parsed(shape=SAMPLE_SHAPE):
    """The parsed workbook whose (clients, holdings) match `shape`."""
    for path in sorted(REAL_DATA_DIR.glob("*.xlsx")):
        parsed = parse_dashboard_workbook(path)
        if (len(parsed.clients), sum(len(c.holdings) for c in parsed.clients)) == shape:
            return parsed
    pytest.skip(f"no client file present parsing to {shape[0]} clients / {shape[1]} holdings")


def _at_on_step(step: int, result, **state) -> AppTest:
    at = AppTest.from_file(str(PAGE), default_timeout=120)
    # These tests exercise screens 2 and 3, which sit behind the access
    # gate. Authenticate up front so each test asserts on the screen it
    # names; the gate itself is covered separately below.
    at.session_state["_authenticated"] = True
    at.session_state["step"] = step
    at.session_state["parse_result"] = result
    at.session_state["parse_error"] = None
    at.session_state["overrides"] = {}
    at.session_state["warnings_ack"] = False
    at.session_state["assembled"] = None
    at.session_state["summary_draft"] = None
    at.session_state["summary_source"] = None
    at.session_state["deliverables"] = None
    for key, value in state.items():
        at.session_state[key] = value
    return at.run()


def _button(at, label_fragment):
    for b in at.button:
        if label_fragment.lower() in b.label.lower():
            return b
    raise AssertionError(
        f"no button matching {label_fragment!r}; present: {[b.label for b in at.button]}"
    )


# ==========================================================================
# Screen 2 - select client and review
# ==========================================================================

def test_client_dropdown_lists_every_detected_client():
    result = _parsed()
    at = _at_on_step(2, result)
    assert not at.exception, at.exception

    assert at.selectbox, "no client dropdown rendered on screen 2"
    dropdown = at.selectbox[0]
    assert list(dropdown.options) == result.client_names()
    assert len(dropdown.options) == 2, dropdown.options


def test_every_client_in_a_nine_client_file_is_listed():
    """The dropdown must not silently truncate. The nine-client book is
    the one that would expose it."""
    result = _parsed((9, 311))
    at = _at_on_step(2, result)
    assert not at.exception, at.exception
    assert list(at.selectbox[0].options) == result.client_names()
    assert len(at.selectbox[0].options) == 9


def test_proceed_is_disabled_until_warnings_are_acknowledged():
    """A client carrying warnings cannot proceed on a glance."""
    result = _parsed()
    client = next(c for c in result.clients
                  if c.warnings or assemble_report_context(
                      client=c, as_of=date.today(), rm=RM).unclassified_schemes)

    at = _at_on_step(2, result, selected_client=client.name)
    assert not at.exception, at.exception

    proceed = _button(at, "Proceed")
    assert proceed.disabled is True, (
        f"Proceed was enabled for {client.name!r} while warnings were outstanding"
    )
    assert any("acknowledge" in c.value.lower() for c in at.caption), \
        [c.value for c in at.caption]

    # Acknowledge, and only then does it open.
    ack = next(cb for cb in at.checkbox if "reviewed the warnings" in cb.label.lower())
    at = ack.check().run()
    assert not at.exception, at.exception
    assert _button(at, "Proceed").disabled is False, \
        "Proceed stayed disabled after the warnings were acknowledged"


def test_client_with_no_warnings_can_proceed_immediately():
    """The gate is on warnings, not on ceremony - a clean client should
    not need a checkbox tick."""
    result = _parsed()
    clean = [
        c for c in result.clients
        if not c.warnings
        and not assemble_report_context(client=c, as_of=date.today(), rm=RM).unclassified_schemes
    ]
    if not clean:
        pytest.skip("no warning-free client in this file")
    at = _at_on_step(2, result, selected_client=clean[0].name)
    assert _button(at, "Proceed").disabled is False


def test_unclassified_scheme_dropdowns_appear_and_reach_the_report_context():
    """The whole point of screen 2: a build warning becomes a human
    decision, and that decision has to actually land in the report."""
    result = _parsed()
    target = None
    for c in result.clients:
        if assemble_report_context(client=c, as_of=date.today(), rm=RM).unclassified_schemes:
            target = c
            break
    if target is None:
        pytest.skip("no unclassified schemes in this file")

    unclassified = assemble_report_context(
        client=target, as_of=date.today(), rm=RM).unclassified_schemes

    at = _at_on_step(2, result, selected_client=target.name)
    assert not at.exception, at.exception

    # One dropdown per unclassified scheme, keyed by scheme name.
    override_boxes = {s.key: s for s in at.selectbox if s.key and s.key.startswith("override::")}
    assert len(override_boxes) == len(unclassified), (
        f"expected a dropdown for each of {unclassified}, got {list(override_boxes)}"
    )
    for scheme in unclassified:
        assert f"override::{scheme}" in override_boxes

    # Choose a real category for the first one, acknowledge, proceed.
    scheme = unclassified[0]
    box = override_boxes[f"override::{scheme}"]
    choice = next(o for o in box.options if o != "Unclassified")
    at = box.select(choice).run()
    ack = next(cb for cb in at.checkbox if "reviewed the warnings" in cb.label.lower())
    at = ack.check().run()
    at = _button(at, "Proceed").click().run()
    assert not at.exception, at.exception

    # The decision reached the assembled report context.
    assert at.session_state["step"] == 3
    assembled = at.session_state["assembled"]
    assert assembled is not None
    rows = {r.label for r in assembled.ctx.equity_sub_allocation}
    assert choice in rows, (
        f"RM assigned {scheme!r} -> {choice!r}, but the Equity Sub-Allocation table has {rows}"
    )


# ==========================================================================
# Screen 3 - summary approval and generate
# ==========================================================================

def _smallest_client(result):
    """The client with the fewest holdings - the fastest one to render.

    Chosen by size rather than by name so no real client name is written
    into this file; these are live client records.
    """
    return min(result.clients, key=lambda c: len(c.holdings))


def _at_on_step3(result, client_name=None):
    client = result.client(client_name) if client_name else result.clients[0]
    assembled = assemble_report_context(client=client, as_of=date.today(), rm=RM)
    return _at_on_step(3, result, selected_client=client.name, assembled=assembled)


def test_generate_is_disabled_until_the_approval_checkbox_is_ticked():
    """The gate. Unticked means no report, and the reason is on screen."""
    at = _at_on_step3(_parsed())
    assert not at.exception, at.exception

    approve = next(cb for cb in at.checkbox if "approve this summary" in cb.label.lower())
    assert approve.value is False, "the approval checkbox must default to unticked"
    assert _button(at, "Generate report").disabled is True
    assert any("approve the summary" in c.value.lower() for c in at.caption), \
        [c.value for c in at.caption]

    at = approve.check().run()
    assert not at.exception, at.exception
    assert _button(at, "Generate report").disabled is False


def test_draft_is_labelled_as_ai_generated_and_is_editable():
    at = _at_on_step3(_parsed())
    assert at.text_area, "the summary must be editable, not fixed text"
    body = " ".join(m.value for m in at.markdown).lower()
    assert "ai-generated draft" in body and "requires review" in body, \
        "the draft must be marked as AI-generated and unreviewed"


@pytest.mark.render
def test_approving_generates_both_deliverables_and_sets_the_gate_flags():
    """End to end through the real page: tick approve, click generate,
    and assert BOTH the flags that protect a client and the two download
    buttons that are the actual output.

    Marked `render` - it builds a real PDF through LibreOffice.
    """
    result = _parsed()
    at = _at_on_step3(result, _smallest_client(result).name)

    edited = "Edited by the RM before approval."
    at.text_area[0].set_value(edited).run()
    approve = next(cb for cb in at.checkbox if "approve this summary" in cb.label.lower())
    at = approve.check().run()
    at = _button(at, "Generate report").click().run()
    assert not at.exception, at.exception

    ctx = at.session_state["assembled"].ctx
    assert ctx.client_summary is not None, "no summary was attached to the report context"
    assert ctx.client_summary.approved is True, "approval did not set approved=True"
    assert ctx.allow_missing_summary is False, \
        "allow_missing_summary must be False once a summary is mandatory"
    assert ctx.client_summary.text == edited, "the RM's edit did not reach the report"

    deliverables = at.session_state["deliverables"]
    assert deliverables is not None, "generation produced no deliverables"
    pdf_path, docx_path = deliverables
    assert Path(pdf_path).exists() and Path(docx_path).exists()

    labels = [d.label for d in at.download_button]
    assert len(labels) == 2, f"expected a PDF and a .docx download, got {labels}"
    assert any("pdf" in l.lower() for l in labels), labels
    assert any("docx" in l.lower() or "word" in l.lower() for l in labels), labels


def test_unapproved_context_would_be_refused_by_the_builder():
    """Belt and braces: even if the UI gate were bypassed, build_report
    refuses an unapproved summary. The two guards are independent."""
    from pipeline.docx_builder import build_report
    from pipeline.summary_client import ClientSummary
    import tempfile

    result = _parsed()
    ctx = assemble_report_context(client=result.clients[0], as_of=date.today(), rm=RM).ctx
    ctx.client_summary = ClientSummary(text="Draft.", approved=False)
    with tempfile.TemporaryDirectory() as tmp:
        with pytest.raises(ValueError, match="has not been approved"):
            build_report(ctx, Path(tmp) / "blocked.docx")


# ==========================================================================
# Streamlit Community Cloud has no LibreOffice
# ==========================================================================

def test_docx_still_ships_when_libreoffice_is_absent(tmp_path, monkeypatch):
    """Without LibreOffice there is no PDF and no way to resolve the TOC -
    but the .docx is built BEFORE conversion is attempted and is complete.
    Losing it as well would turn one missing dependency into no deliverable
    at all, which is what happened before this path existed.
    """
    import pipeline.pdf_converter as pc

    monkeypatch.setattr(pc, "_find_soffice",
                        lambda: (_ for _ in ()).throw(FileNotFoundError("no soffice")))
    assert pc.soffice_available() is False

    result = _parsed()
    assembled = assemble_report_context(
        client=_smallest_client(result), as_of=date.today(), rm=RM)
    from pipeline.summary_client import ClientSummary
    assembled.ctx.client_summary = ClientSummary(text="Approved.", approved=True)
    assembled.ctx.allow_missing_summary = False

    with pytest.raises(pc.PdfUnavailable) as excinfo:
        pc.build_report_deliverables(assembled.ctx, tmp_path)

    docx_path = excinfo.value.docx_path
    assert docx_path.exists(), "the .docx must survive a missing PDF renderer"
    assert docx_path.stat().st_size > 10_000
    assert "LibreOffice" in str(excinfo.value)

    # And the non-raising form returns it directly.
    pdf, docx, pages = pc.build_report_deliverables(
        assembled.ctx, tmp_path / "b", require_pdf=False)
    assert pdf is None and docx.exists() and pages == {}


def test_app_offers_the_docx_when_pdf_rendering_is_unavailable(monkeypatch):
    """End to end through the page: approve, generate, and confirm the RM
    is warned and still gets a download."""
    import pipeline.pdf_converter as pc
    monkeypatch.setattr(pc, "_find_soffice",
                        lambda: (_ for _ in ()).throw(FileNotFoundError("no soffice")))

    result = _parsed()
    at = _at_on_step3(result, _smallest_client(result).name)
    warnings_before = " ".join(w.value for w in at.warning).lower()
    assert "pdf rendering is unavailable" in warnings_before, \
        "the RM must be told before approving, not after generating"

    approve = next(cb for cb in at.checkbox if "approve this summary" in cb.label.lower())
    at = approve.check().run()
    at = _button(at, "Generate report").click().run()
    assert not at.exception, at.exception
    assert not at.error, [e.value[:200] for e in at.error]

    pdf_path, docx_path = at.session_state["deliverables"]
    assert pdf_path is None
    assert Path(docx_path).exists()

    labels = [d.label for d in at.download_button]
    assert len(labels) == 1 and "docx" in labels[0].lower(), labels


# ==========================================================================
# Access gate
# ==========================================================================

def test_app_refuses_access_when_no_password_is_configured(monkeypatch):
    """Fails CLOSED. A gate that disables itself when misconfigured is
    worse than no gate, because it looks like protection."""
    import pipeline.app_secrets as secrets
    monkeypatch.setattr(secrets, "get_app_password", lambda: None)

    at = AppTest.from_file(str(PAGE), default_timeout=60).run()
    assert not at.exception, at.exception
    assert any("not configured for access" in e.value.lower() for e in at.error), \
        [e.value for e in at.error]
    # ...and nothing past the gate rendered.
    assert not at.file_uploader, "the upload widget rendered without a password"


def test_password_prompt_blocks_the_page_until_matched(monkeypatch):
    import pipeline.app_secrets as secrets
    monkeypatch.setattr(secrets, "get_app_password", lambda: "correct-horse")

    at = AppTest.from_file(str(PAGE), default_timeout=60).run()
    assert not at.exception, at.exception
    assert at.text_input, "no password prompt rendered"
    # proto.type 1 == PASSWORD; .type is the element name, not the input mode.
    assert at.text_input[0].proto.type == 1, "the password field must be masked"
    assert not at.file_uploader, "page content rendered before authentication"

    # Wrong password: refused, still gated.
    at = at.text_input[0].set_value("wrong").run()
    assert any("incorrect password" in e.value.lower() for e in at.error)
    assert not at.file_uploader

    # Correct password: through, and the flag is stored.
    at = at.text_input[0].set_value("correct-horse").run()
    assert not at.error, [e.value for e in at.error]
    assert at.session_state["_authenticated"] is True
    assert at.file_uploader, "authenticated page did not render its content"


def test_authentication_is_not_re_prompted_on_rerun(monkeypatch):
    """Every widget interaction is a rerun; re-prompting each time would
    make the app unusable."""
    import pipeline.app_secrets as secrets
    monkeypatch.setattr(secrets, "get_app_password", lambda: "correct-horse")

    at = AppTest.from_file(str(PAGE), default_timeout=60)
    at.session_state["_authenticated"] = True
    at.run()
    assert not at.exception, at.exception
    assert not at.text_input, "already-authenticated session was prompted again"
    assert at.file_uploader


def test_password_is_not_left_in_session_state(monkeypatch):
    """The entered password should not sit in session state after use."""
    import pipeline.app_secrets as secrets
    monkeypatch.setattr(secrets, "get_app_password", lambda: "correct-horse")
    at = AppTest.from_file(str(PAGE), default_timeout=60).run()
    at = at.text_input[0].set_value("correct-horse").run()
    assert at.session_state["_authenticated"] is True
    held = at.session_state["_password_input"] if "_password_input" in at.session_state else None
    assert not held, "the plaintext password is still held in session state"


# ==========================================================================
# cairosvg / libcairo missing
# ==========================================================================

class _BlockCairosvg:
    """Simulates a Debian image without libcairo2: importing cairosvg
    raises OSError from cffi's dlopen, not ImportError."""

    def find_module(self, name, path=None):
        if name == "cairosvg":
            raise OSError("simulated: cannot load library 'libcairo.so.2'")
        return None

    def find_spec(self, name, path=None, target=None):
        if name == "cairosvg":
            raise OSError("simulated: cannot load library 'libcairo.so.2'")
        return None


@pytest.fixture
def no_cairosvg(monkeypatch):
    import sys
    blocker = _BlockCairosvg()
    monkeypatch.setitem(sys.modules, "cairosvg", None)
    monkeypatch.delitem(sys.modules, "cairosvg", raising=False)
    sys.meta_path.insert(0, blocker)
    try:
        yield
    finally:
        sys.meta_path.remove(blocker)


def test_mindmap_module_imports_without_cairosvg(no_cairosvg):
    """The regression that took down a whole report: cairosvg was
    imported at module level, so a client with zero transactions - who
    has no mind map at all - still hit the missing system library."""
    import importlib
    import pipeline.mindmap as mm
    importlib.reload(mm)
    assert mm.cairosvg_available() is False
    # The pure-Python half must remain usable.
    assert mm.build_mindmap_recommendations_from_transactions([]) == []


def test_report_generates_for_a_client_with_no_actions_without_cairosvg(no_cairosvg, tmp_path):
    """JITENDER's case exactly: no actions, no mind map needed, and a
    missing renderer must not block the report."""
    from pipeline.report_assembler import assemble_report_context
    from pipeline.docx_builder import build_report
    from pipeline.summary_client import ClientSummary

    result = _parsed((3, 65))          # the book with zero actions
    client = _smallest_client(result)
    assert not client.actions, "this fixture is expected to carry no actions"

    assembled = assemble_report_context(client=client, as_of=date.today(), rm=RM)
    assembled.ctx.client_summary = ClientSummary(text="Approved.", approved=True)
    out = build_report(assembled.ctx, tmp_path / "no_actions.docx")
    assert out.exists() and out.stat().st_size > 10_000


def test_report_generates_with_recommendations_but_no_renderer(no_cairosvg, tmp_path):
    """A client who DOES have recommendations still gets a report - the
    diagram is skipped, the recommendations survive in the Transaction
    Snapshot, and the section says which of the two happened."""
    from pipeline.mindmap import MindmapUnavailable, generate_mindmap
    from pipeline.report_assembler import assemble_report_context
    from pipeline.docx_builder import build_report
    from pipeline.summary_client import ClientSummary
    from docx import Document

    result = _parsed()
    client = next(c for c in result.clients if c.actions)
    assembled = assemble_report_context(client=client, as_of=date.today(), rm=RM)
    assert assembled.mindmap_recommendations, "expected recommendations to draw"

    with pytest.raises(MindmapUnavailable) as excinfo:
        generate_mindmap(assembled.mindmap_recommendations,
                         client_name=client.name, output_path=tmp_path / "mm.png")
    assert "libcairo2" in str(excinfo.value)

    ctx = assembled.ctx
    ctx.mindmap_path = None
    ctx.mindmap_unavailable_reason = str(excinfo.value)
    ctx.client_summary = ClientSummary(text="Approved.", approved=True)

    out = build_report(ctx, tmp_path / "with_recs.docx")
    assert out.exists()

    body = "\n".join(p.text for p in Document(str(out)).paragraphs)
    assert "could not be rendered" in body, \
        "the report must say the diagram is missing"
    assert "No recommended changes for this review cycle." not in body, \
        "a missing renderer must never be reported as the client having no recommendations"


def test_app_generates_report_when_cairosvg_is_unavailable(no_cairosvg):
    """End to end through the page with the renderer gone."""
    result = _parsed()
    at = _at_on_step3(result, next(c for c in result.clients if c.actions).name)
    approve = next(cb for cb in at.checkbox if "approve this summary" in cb.label.lower())
    at = approve.check().run()
    at = _button(at, "Generate report").click().run()

    assert not at.exception, at.exception
    assert not at.error, [e.value[:300] for e in at.error]
    assert at.session_state["deliverables"] is not None, \
        "a missing optional renderer blocked the whole report"
    assert any("mind map could not be drawn" in w.value.lower() for w in at.warning), \
        [w.value for w in at.warning]
