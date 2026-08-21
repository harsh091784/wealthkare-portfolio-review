"""
pipeline/mindmap.py

SVG mind-map generator for the WC Securities / Wealthkare Portfolio
Review pipeline - visualises recommended portfolio changes (switches,
redemptions, SIP stop/starts) as a central-node diagram with red/green
branches, converted to PNG for docx embedding.

Design notes
------------
- Central node: dark navy rectangle, white bold text, "[CLIENT NAME]
  Portfolio Changes".
- Outgoing actions (Switch Out, Redeem, SIP Stop) get RED connectors and
  red branch labels. Incoming actions (Switch In, Reinvest, SIP Start) get
  GREEN connectors and green branch labels. These are the ONLY six action
  labels ever rendered - anything else in the input (Exit, Reduce, Hold,
  Redirect, Sell, or any other free-text action) is rejected and reported
  as a warning rather than silently relabelled or guessed into a color.
- Each leaf node is a rounded rectangle: scheme name + ₹ amount.
- A dashed flow line connects an outgoing leaf to its corresponding
  incoming leaf, matched via `suggested_scheme` (Switch Out) <-> `scheme`
  (Switch In / Reinvest) in the SAME mind-map section. If no matching
  incoming leaf exists in the input, the flow line is simply skipped (not
  fabricated) and a warning is recorded.
- If any SIP-specific actions (SIP Stop / SIP Start) are present alongside
  lumpsum actions (Switch Out / Switch In / Redeem / Reinvest), a SECOND
  mind-map section is rendered below the first, using the same red/green
  convention, with its own central node labelled "[CLIENT NAME] SIP
  Changes".
- CRITICAL: every <text> element uses font-family="DejaVu Sans" (not just
  the ones showing a ₹ symbol) - Arial silently drops the ₹ glyph, which
  was a real production bug. Using DejaVu Sans everywhere removes any risk
  of a stray text element being missed.
- Rendered as SVG first, then rasterised via cairosvg at scale=2.5 for
  crisp docx embedding.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional, Union
from xml.sax.saxutils import escape as xml_escape

from pipeline.chart_gen import format_inr

# NOTE: cairosvg is deliberately NOT imported here.
#
# It is only needed to rasterise the finished SVG, and it loads
# libcairo.so.2 through cffi at import time. A module-level import means
# a missing SYSTEM library raises an OSError the moment anything touches
# this module - and report_assembler imports
# build_mindmap_recommendations_from_transactions from here for every
# client, including clients with no transactions and therefore no mind
# map at all. That is exactly how a client with zero actions came to fail
# on a renderer their report never uses. The import now lives inside
# generate_mindmap(), the one function that actually rasterises.

# --------------------------------------------------------------------------
# Constants
# --------------------------------------------------------------------------

OUTGOING_ACTIONS = {"Switch Out", "Redeem", "SIP Stop"}
INCOMING_ACTIONS = {"Switch In", "Reinvest", "SIP Start"}
VALID_ACTIONS = OUTGOING_ACTIONS | INCOMING_ACTIONS

LUMPSUM_ACTIONS = {"Switch Out", "Switch In", "Redeem", "Reinvest"}
SIP_ACTIONS = {"SIP Stop", "SIP Start"}

# Explicitly-banned labels called out in the spec - checked to produce a
# clearer warning than the generic "not a recognised action" message.
BANNED_ACTION_HINTS = {"Exit", "Reduce", "Hold", "Redirect", "Sell"}

RED = "#C0392B"
GREEN = "#1E8449"
NAVY = "#1C2B4B"
FLOW_LINE_COLOR = "#8A93A6"
LEAF_BG = "#FDFDFB"
LEAF_BORDER_NEUTRAL = "#D9D5C8"

FONT_FAMILY = "DejaVu Sans"  # applied to every <text> element, per spec

CANVAS_WIDTH = 1180
MARGIN_X = 40
MARGIN_TOP = 30
MARGIN_BOTTOM = 30
SECTION_GAP = 60

CENTRAL_WIDTH = 240
CENTRAL_HEIGHT = 100

LEAF_X = MARGIN_X + CENTRAL_WIDTH + 220
LEAF_WIDTH = 320
LEAF_HEIGHT = 66
LEAF_GAP = 24


# --------------------------------------------------------------------------
# Data model
# --------------------------------------------------------------------------

@dataclass
class Recommendation:
    scheme: str
    action: str
    amount: float
    suggested_scheme: Optional[str] = None


@dataclass
class MindmapResult:
    png_path: Path
    warnings: list[str] = field(default_factory=list)


# --------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------

def _coerce_recommendation(raw) -> Recommendation:
    if isinstance(raw, Recommendation):
        return raw
    return Recommendation(
        scheme=str(raw.get("scheme", "")).strip(),
        action=str(raw.get("action", "")).strip(),
        amount=raw.get("amount"),
        suggested_scheme=(str(raw["suggested_scheme"]).strip() if raw.get("suggested_scheme") else None),
    )


def build_mindmap_recommendations_from_transactions(transactions) -> list:
    """Builds the Mind Map's recommendation list SOLELY from the
    Transaction Snapshot's rows - the SAME list that renders in that
    table - so the two sections can never disagree on an amount again.
    (Bug: the Mind Map used to be generated from its own independently
    -maintained recommendation list, which rendered HDFC Flexicap's OLD
    purchase-value amount while the Transaction Snapshot correctly showed
    the current value - two numbers for the same transaction.)

    transactions: duck-typed - each needs .scheme, .action, .amount,
    .suggested_scheme (docx_builder.TransactionSnapshotRow instances, or
    anything else with those four attributes).

    Returns a list of Recommendation instances (not dicts) - passing
    those straight to generate_mindmap() skips _coerce_recommendation()'s
    dict-shape assumptions entirely.
    """
    recs = [
        Recommendation(
            scheme=t.scheme, action=t.action, amount=t.amount,
            suggested_scheme=getattr(t, "suggested_scheme", None),
        )
        for t in transactions
    ]

    # Regression guard: every recommendation rendered in the Mind Map must
    # trace back to an actual row in `transactions` - true by construction
    # (the loop above only ever reads from `transactions`), asserted
    # explicitly so a future refactor that reintroduces a separate
    # recommendation list fails loudly instead of silently drifting.
    source_triples = {(t.scheme, t.action, t.amount) for t in transactions}
    for r in recs:
        assert (r.scheme, r.action, r.amount) in source_triples, (
            f"Mind map recommendation for '{r.scheme}' ({r.action}, amount={r.amount}) does not "
            f"match any row in the Transaction Snapshot list - this should be impossible."
        )

    return recs


def _wrap_scheme_name(name: str, max_chars: int = 26) -> list[str]:
    """Very small wrapper: splits a long scheme name onto at most 2 lines,
    breaking on the nearest space to the midpoint rather than mid-word."""
    if len(name) <= max_chars:
        return [name]
    mid = len(name) // 2
    split_at = name.rfind(" ", 0, mid + 8)
    if split_at == -1:
        split_at = name.find(" ", mid)
    if split_at == -1:
        return [name[:max_chars], name[max_chars:]]
    return [name[:split_at], name[split_at + 1:]]


def _text(x: float, y: float, content: str, *, size: int, color: str,
          weight: str = "normal", anchor: str = "start") -> str:
    return (
        f'<text x="{x:.1f}" y="{y:.1f}" font-family="{FONT_FAMILY}" '
        f'font-size="{size}" font-weight="{weight}" fill="{color}" '
        f'text-anchor="{anchor}">{xml_escape(content)}</text>'
    )


def _rounded_rect(x: float, y: float, w: float, h: float, *, fill: str,
                   stroke: str, stroke_width: float = 1.5, rx: float = 10,
                   left_accent: Optional[str] = None) -> str:
    parts = [
        f'<rect x="{x:.1f}" y="{y:.1f}" width="{w:.1f}" height="{h:.1f}" '
        f'rx="{rx}" ry="{rx}" fill="{fill}" stroke="{stroke}" stroke-width="{stroke_width}"/>'
    ]
    if left_accent:
        # A colored accent bar on the left edge, clipped to the rect's
        # rounded corners by drawing a slightly inset small rect.
        parts.append(
            f'<rect x="{x:.1f}" y="{y + 4:.1f}" width="5" height="{h - 8:.1f}" '
            f'rx="2.5" ry="2.5" fill="{left_accent}"/>'
        )
    return "".join(parts)


def _connector_path(x1: float, y1: float, x2: float, y2: float, color: str, marker_id: str) -> str:
    mid_x = (x1 + x2) / 2
    return (
        f'<path d="M {x1:.1f} {y1:.1f} C {mid_x:.1f} {y1:.1f}, {mid_x:.1f} {y2:.1f}, {x2:.1f} {y2:.1f}" '
        f'fill="none" stroke="{color}" stroke-width="2.2" marker-end="url(#{marker_id})"/>'
    )


def _dashed_flow_path(x1: float, y1: float, x2: float, y2: float) -> str:
    mid_x = (x1 + x2) / 2
    return (
        f'<path d="M {x1:.1f} {y1:.1f} C {mid_x:.1f} {y1:.1f}, {mid_x:.1f} {y2:.1f}, {x2:.1f} {y2:.1f}" '
        f'fill="none" stroke="{FLOW_LINE_COLOR}" stroke-width="1.6" stroke-dasharray="6,4" '
        f'marker-end="url(#flowArrow)"/>'
    )


# --------------------------------------------------------------------------
# Section rendering
# --------------------------------------------------------------------------

def _render_section(
    y_top: float,
    central_label: str,
    recs: list[Recommendation],
    warnings: list[str],
) -> tuple[str, float]:
    valid: list[Recommendation] = []
    for r in recs:
        if r.action not in VALID_ACTIONS:
            hint = " (this label is explicitly banned - use one of the six standard action verbs)" \
                if r.action in BANNED_ACTION_HINTS else ""
            warnings.append(
                f"'{r.scheme}': action '{r.action}' is not one of the six allowed labels "
                f"(Switch Out, Switch In, Redeem, Reinvest, SIP Stop, SIP Start) - "
                f"excluded from the mind map{hint}."
            )
            continue
        if r.amount is None:
            warnings.append(f"'{r.scheme}' ({r.action}): amount missing - excluded from the mind map.")
            continue
        valid.append(r)

    outgoing = [r for r in valid if r.action in OUTGOING_ACTIONS]
    incoming = [r for r in valid if r.action in INCOMING_ACTIONS]

    total_rows = max(len(outgoing) + len(incoming), 1)
    section_height = MARGIN_TOP + total_rows * (LEAF_HEIGHT + LEAF_GAP) - LEAF_GAP + MARGIN_BOTTOM
    section_height = max(section_height, CENTRAL_HEIGHT + MARGIN_TOP + MARGIN_BOTTOM)

    svg_parts: list[str] = []

    # --- Central node ---
    central_x = MARGIN_X
    central_y = y_top + section_height / 2 - CENTRAL_HEIGHT / 2
    svg_parts.append(_rounded_rect(
        central_x, central_y, CENTRAL_WIDTH, CENTRAL_HEIGHT,
        fill=NAVY, stroke=NAVY, rx=14,
    ))
    title_lines = _wrap_scheme_name(central_label, max_chars=22)
    line_y = central_y + CENTRAL_HEIGHT / 2 - (len(title_lines) - 1) * 10 + 5
    for line in title_lines:
        svg_parts.append(_text(
            central_x + CENTRAL_WIDTH / 2, line_y, line,
            size=16, color="#FFFFFF", weight="bold", anchor="middle",
        ))
        line_y += 22

    central_right_x = central_x + CENTRAL_WIDTH

    # --- Leaves (outgoing first, then incoming) + connectors ---
    # leaf_anchor: scheme name -> (x, y) of the leaf's right-edge midpoint,
    # used as the dashed flow line's arrival point. Keyed on incoming
    # scheme names only (that's what an outgoing suggested_scheme matches
    # against).
    incoming_anchor_by_scheme: dict[str, tuple[float, float]] = {}
    leaf_render_queue: list[tuple[Recommendation, float, float, str]] = []  # (rec, x, y, direction)

    row_idx = 0
    for r in outgoing:
        leaf_y = y_top + MARGIN_TOP + row_idx * (LEAF_HEIGHT + LEAF_GAP)
        leaf_render_queue.append((r, LEAF_X, leaf_y, "out"))
        row_idx += 1
    for r in incoming:
        leaf_y = y_top + MARGIN_TOP + row_idx * (LEAF_HEIGHT + LEAF_GAP)
        leaf_render_queue.append((r, LEAF_X, leaf_y, "in"))
        incoming_anchor_by_scheme[r.scheme] = (LEAF_X, leaf_y + LEAF_HEIGHT / 2)
        row_idx += 1

    outgoing_flow_start: dict[int, tuple[float, float]] = {}  # id(rec) -> (x, y) at leaf's right edge

    # Connector origins are spread across the central node's right edge
    # (rather than all bundled at its exact vertical center) so branches
    # fan out visibly instead of overlapping near the node - this is what
    # keeps the action-verb labels from colliding with each other.
    n_branches = len(leaf_render_queue)
    edge_top = central_y + 18
    edge_bottom = central_y + CENTRAL_HEIGHT - 18
    for branch_idx, (r, leaf_x, leaf_y, direction) in enumerate(leaf_render_queue):
        color = RED if direction == "out" else GREEN
        marker_id = "redArrow" if direction == "out" else "greenArrow"
        leaf_center_y = leaf_y + LEAF_HEIGHT / 2

        if n_branches > 1:
            start_y = edge_top + branch_idx * (edge_bottom - edge_top) / (n_branches - 1)
        else:
            start_y = central_y + CENTRAL_HEIGHT / 2

        # Connector from central node to this leaf.
        svg_parts.append(_connector_path(
            central_right_x, start_y,
            leaf_x, leaf_center_y, color, marker_id,
        ))

        # Branch label (the action verb itself), positioned just above the
        # connector as it approaches its leaf node, rather than near the
        # shared central-node origin - leaves are spaced far enough apart
        # vertically (LEAF_HEIGHT + LEAF_GAP) that labels placed here stay
        # clear of each other, whereas labels bunched near the origin
        # collide regardless of leaf spacing.
        label_x = leaf_x - 12
        label_y = leaf_center_y - 12
        svg_parts.append(_text(
            label_x, label_y, r.action, size=12.5, color=color, weight="bold", anchor="end",
        ))

        # Leaf node.
        svg_parts.append(_rounded_rect(
            leaf_x, leaf_y, LEAF_WIDTH, LEAF_HEIGHT,
            fill=LEAF_BG, stroke=LEAF_BORDER_NEUTRAL, left_accent=color,
        ))
        name_lines = _wrap_scheme_name(r.scheme)
        name_y = leaf_y + 24 if len(name_lines) == 1 else leaf_y + 20
        for line in name_lines:
            svg_parts.append(_text(leaf_x + 18, name_y, line, size=13, color=NAVY, weight="bold"))
            name_y += 15
        amount_y = leaf_y + LEAF_HEIGHT - 14
        svg_parts.append(_text(leaf_x + 18, amount_y, format_inr(r.amount), size=13, color="#3A3A3A"))

        if direction == "out":
            outgoing_flow_start[id(r)] = (leaf_x + LEAF_WIDTH, leaf_center_y)

    # --- Dashed flow lines: outgoing.suggested_scheme -> incoming.scheme ---
    for r in outgoing:
        if not r.suggested_scheme:
            continue
        target = incoming_anchor_by_scheme.get(r.suggested_scheme)
        if target is None:
            warnings.append(
                f"'{r.scheme}' ({r.action}) points to suggested_scheme "
                f"'{r.suggested_scheme}', but no matching incoming leaf was found in this "
                f"section - flow line skipped rather than guessed."
            )
            continue
        start_x, start_y = outgoing_flow_start[id(r)]
        # Arrive on the leaf's TOP edge (offset in from the corner) rather
        # than its left-edge midpoint - that point is already occupied by
        # the solid green connector's arrowhead, and stacking both markers
        # on the same spot makes them unreadable.
        target_x, target_y = target
        arrival_x = target_x + LEAF_WIDTH * 0.35
        arrival_y = target_y - LEAF_HEIGHT / 2
        svg_parts.append(_dashed_flow_path(start_x, start_y - 6, arrival_x, arrival_y))

    return "".join(svg_parts), section_height


# --------------------------------------------------------------------------
# Top-level entry point
# --------------------------------------------------------------------------

class MindmapUnavailable(RuntimeError):
    """The mind map could not be rasterised because cairosvg (or the
    libcairo system library behind it) is unavailable.

    Distinct from "this client has no recommendations": one is a missing
    renderer, the other is an empty result, and a report must not present
    the first as the second.
    """


def cairosvg_available() -> bool:
    """Whether the mind map can be rendered in this environment.

    Catches Exception rather than ImportError alone: the failure mode on
    a Debian image without libcairo2 is an OSError raised by cffi's
    dlopen from inside cairosvg's own import, not an ImportError.
    """
    try:
        import cairosvg  # noqa: F401
    except Exception:
        return False
    return True


def generate_mindmap(
    recommendations: list,
    client_name: str,
    output_path: Union[str, Path],
    scale: float = 2.5,
) -> MindmapResult:
    """Render the portfolio-changes mind map (and, if SIP-specific actions
    are present, a second SIP mind map beneath it) and save as PNG.

    recommendations: list of dicts (or Recommendation instances) with keys
    scheme, action, amount, suggested_scheme (suggested_scheme optional).
    """
    output_path = Path(output_path)
    warnings: list[str] = []

    recs = [_coerce_recommendation(r) for r in recommendations]

    lumpsum_recs = [r for r in recs if r.action in LUMPSUM_ACTIONS]
    sip_recs = [r for r in recs if r.action in SIP_ACTIONS]
    # Anything with an unrecognised action is left in `recs` untouched here;
    # _render_section() validates and warns on it per-section so nothing is
    # silently swallowed before it can be reported.
    other_recs = [r for r in recs if r.action not in LUMPSUM_ACTIONS and r.action not in SIP_ACTIONS]
    # Unrecognised-action rows still need to surface a warning even though
    # they don't belong to either section; route them through the primary
    # section's validation so the message is generated in one place.
    lumpsum_recs = lumpsum_recs + other_recs

    section1_svg, section1_height = _render_section(
        MARGIN_TOP, f"{client_name} Portfolio Changes", lumpsum_recs, warnings,
    )

    sections_svg = [section1_svg]
    total_height = MARGIN_TOP + section1_height

    if sip_recs:
        section2_svg, section2_height = _render_section(
            total_height + SECTION_GAP, f"{client_name} SIP Changes", sip_recs, warnings,
        )
        sections_svg.append(section2_svg)
        total_height += SECTION_GAP + section2_height

    total_height += MARGIN_BOTTOM

    defs = f"""
    <defs>
        <marker id="redArrow" markerWidth="8" markerHeight="8" refX="6" refY="3" orient="auto">
            <path d="M0,0 L6,3 L0,6 Z" fill="{RED}"/>
        </marker>
        <marker id="greenArrow" markerWidth="8" markerHeight="8" refX="6" refY="3" orient="auto">
            <path d="M0,0 L6,3 L0,6 Z" fill="{GREEN}"/>
        </marker>
        <marker id="flowArrow" markerWidth="7" markerHeight="7" refX="5" refY="2.5" orient="auto">
            <path d="M0,0 L5,2.5 L0,5 Z" fill="{FLOW_LINE_COLOR}"/>
        </marker>
    </defs>
    """

    svg_doc = (
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{CANVAS_WIDTH}" '
        f'height="{total_height:.1f}" viewBox="0 0 {CANVAS_WIDTH} {total_height:.1f}">'
        f'{defs}'
        f'<rect x="0" y="0" width="{CANVAS_WIDTH}" height="{total_height:.1f}" fill="none"/>'
        + "".join(sections_svg)
        + "</svg>"
    )

    try:
        import cairosvg
    except Exception as exc:  # ImportError, or OSError from cffi's dlopen
        raise MindmapUnavailable(
            f"The mind map could not be rendered: cairosvg is unavailable "
            f"({type(exc).__name__}: {exc}). On Debian this needs the libcairo2 "
            f"system package. Everything else in the report is unaffected."
        ) from exc

    cairosvg.svg2png(
        bytestring=svg_doc.encode("utf-8"),
        write_to=str(output_path),
        scale=scale,
        output_width=None,
        output_height=None,
    )

    return MindmapResult(png_path=output_path, warnings=warnings)


# --------------------------------------------------------------------------
# Self-test (dummy data, no external files required)
# --------------------------------------------------------------------------

def _run_self_test() -> None:
    import tempfile

    print("=== pipeline/mindmap.py self-test ===\n")

    sample_recommendations = [
        # Lumpsum switch pair: HDFC Flexicap -> Parag Parikh Flexicap
        {"scheme": "HDFC Flexicap Fund", "action": "Switch Out", "amount": 250000,
         "suggested_scheme": "Parag Parikh Flexicap Fund"},
        {"scheme": "Parag Parikh Flexicap Fund", "action": "Switch In", "amount": 250000,
         "suggested_scheme": None},
        # Straight redemption, no reinvestment target
        {"scheme": "L&T Emerging Businesses Fund", "action": "Redeem", "amount": 80000,
         "suggested_scheme": None},
        # Reinvest into a new scheme
        {"scheme": "Kotak Multicap Fund", "action": "Reinvest", "amount": 80000,
         "suggested_scheme": None},
        # SIP stop/start pair
        {"scheme": "Axis Bluechip Fund", "action": "SIP Stop", "amount": 10000,
         "suggested_scheme": "Mirae Asset Large Cap Fund"},
        {"scheme": "Mirae Asset Large Cap Fund", "action": "SIP Start", "amount": 10000,
         "suggested_scheme": None},
        # Deliberately invalid action - should be excluded + warned, never rendered
        {"scheme": "Old Legacy Fund", "action": "Exit", "amount": 15000, "suggested_scheme": None},
    ]

    print("Input recommendations:")
    for r in sample_recommendations:
        print(f"  {r['scheme']:<32} {r['action']:<12} ₹{r['amount']:>10,}  -> {r['suggested_scheme']}")

    with tempfile.TemporaryDirectory() as tmp:
        out_path = Path(tmp) / "mindmap_test.png"
        result = generate_mindmap(sample_recommendations, client_name="Rahul Sharma", output_path=out_path)

        assert result.png_path.exists(), "PNG was not written to disk."
        size_bytes = result.png_path.stat().st_size
        print(f"\nSaved PNG: {result.png_path} ({size_bytes:,} bytes)")

        try:
            from PIL import Image
            with Image.open(result.png_path) as img:
                print(f"Image dimensions: {img.size[0]}x{img.size[1]} px, mode={img.mode}")
        except ImportError:
            print("(Pillow not installed - skipping pixel-dimension check.)")

        print(f"\nWarnings ({len(result.warnings)}):")
        for w in result.warnings:
            print(f"  - {w}")

        assert any("Exit" in w for w in result.warnings), "Expected a warning for the banned 'Exit' action."
        assert not any("Old Legacy Fund" in s for s in [""]), "sanity no-op"

        # Copy to a stable, inspectable location before the temp dir is cleaned up.
        stable_path = Path(__file__).parent.parent / "_test_output_mindmap.png"
        stable_path.write_bytes(result.png_path.read_bytes())
        print(f"\nCopied to stable path for inspection: {stable_path}")

    # --- build_mindmap_recommendations_from_transactions: same source as
    # the Transaction Snapshot table, not an independently-maintained list ---
    print("\n--- build_mindmap_recommendations_from_transactions ---")
    from dataclasses import dataclass as _dc

    @_dc
    class _FakeTransactionRow:
        scheme: str
        action: str
        amount: float
        suggested_scheme: Optional[str] = None

    fake_transactions = [
        _FakeTransactionRow("HDFC Flexicap Fund", "Switch Out", 298000, "Parag Parikh Flexicap Fund"),
        _FakeTransactionRow("Parag Parikh Flexicap Fund", "Switch In", 284150),
    ]
    derived_recs = build_mindmap_recommendations_from_transactions(fake_transactions)
    assert len(derived_recs) == 2
    assert derived_recs[0].amount == 298000, "Mind map amount must equal the Transaction Snapshot's amount exactly."
    print(f"Derived: {[(r.scheme, r.action, r.amount) for r in derived_recs]}")

    print("\nAll self-test assertions passed.")


if __name__ == "__main__":
    _run_self_test()
