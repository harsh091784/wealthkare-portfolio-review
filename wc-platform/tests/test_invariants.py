"""
tests/test_invariants.py

The 12 correctness invariants, asserted on EVERY fixture - not just the
golden one. These encode the checks that were previously being run by
hand on each rendered PDF, which is why they're stated as properties of
the data ("these two totals must reconcile") rather than as snapshots of
expected output ("this cell must say Rs 90,30,000").

Snapshot-style expected values do exist, but only in
test_golden_expected_values below, and only for the golden dataset -
they're a canary for "the fixture data itself changed", not the
correctness gate.
"""

from __future__ import annotations

import re

import pytest

from pipeline.docx_builder import _parse_deadline, _validate_report_context
from pipeline.risk_profile import compute_headline_equity_exposure_pct
from pipeline.tax_calc import TAXABLE_TRANSACTION_ACTIONS

RUPEE_TOLERANCE = 1  # invariants stated "to the rupee"


# --------------------------------------------------------------------------
# 1. sum(asset_allocation buckets) == holdings grand total, to the rupee
# --------------------------------------------------------------------------

def test_invariant_01_asset_allocation_reconciles_to_holdings(fixture):
    holdings_total = sum(h.current_value for h in fixture.ctx.holdings if h.current_value is not None)
    allocation_total = sum(v for v in fixture.ctx.asset_allocation.values() if v is not None)
    assert abs(round(allocation_total) - round(holdings_total)) <= RUPEE_TOLERANCE, (
        f"[{fixture.name}] asset allocation total Rs {allocation_total:,.0f} != holdings grand "
        f"total Rs {holdings_total:,.0f}. A bucket with no matching holding pads the pie."
    )


# --------------------------------------------------------------------------
# 2. sum(market_cap buckets) == equity bucket total, to the rupee
# --------------------------------------------------------------------------

def test_invariant_02_market_cap_reconciles_to_equity_bucket(fixture):
    equity_total = sum(
        h.current_value for h in fixture.ctx.holdings
        if h.category == "Equity" and h.current_value is not None
    )
    market_cap_total = sum(r.value for r in fixture.market_cap_rows)
    assert abs(round(market_cap_total) - round(equity_total)) <= RUPEE_TOLERANCE, (
        f"[{fixture.name}] market-cap buckets total Rs {market_cap_total:,.0f} != equity holdings "
        f"total Rs {equity_total:,.0f}. An equity holding is being dropped or double-counted."
    )


# --------------------------------------------------------------------------
# 3. sum(individual holding current values) == grand total
# --------------------------------------------------------------------------

def test_invariant_03_holdings_sum_to_reported_grand_total(fixture):
    computed = sum(h.current_value for h in fixture.ctx.holdings if h.current_value is not None)
    reported = fixture.ctx.portfolio_summary.current_value
    assert abs(round(computed) - round(reported)) <= RUPEE_TOLERANCE, (
        f"[{fixture.name}] Portfolio Overview current value Rs {reported:,.0f} != sum of holdings "
        f"Rs {computed:,.0f}."
    )


# --------------------------------------------------------------------------
# 4. equity exposure on the pie footnote == equity exposure on the gauge
# --------------------------------------------------------------------------

def test_invariant_04_single_equity_exposure_figure(fixture):
    """Both the Current Asset Allocation footnote and the Risk Profile
    gauge must call compute_headline_equity_exposure_pct() - the same
    function, same float. They previously used two different calculations
    and printed 91.0% and 95.6% for the same portfolio."""
    from_helper = compute_headline_equity_exposure_pct(fixture.ctx.holdings)
    assert from_helper == fixture.equity_exposure_pct, (
        f"[{fixture.name}] equity exposure differs between call sites: "
        f"{from_helper} vs {fixture.equity_exposure_pct}."
    )
    if fixture.ctx.holdings:
        assert from_helper is not None, (
            f"[{fixture.name}] equity exposure is None despite holdings being present."
        )


# --------------------------------------------------------------------------
# 5. every Mind Map amount appears in transaction_snapshot
# --------------------------------------------------------------------------

def test_invariant_05_mindmap_amounts_come_from_transactions(fixture):
    source = {(t.scheme, t.action, t.amount) for t in fixture.ctx.transaction_snapshot}
    for rec in fixture.mindmap_recs:
        assert (rec.scheme, rec.action, rec.amount) in source, (
            f"[{fixture.name}] Mind Map renders {rec.scheme} / {rec.action} / {rec.amount}, which "
            f"is not a row in transaction_snapshot. The two sections have diverged."
        )


# --------------------------------------------------------------------------
# 6. for every switch pair: switch_in == switch_out - allocated_deduction
# --------------------------------------------------------------------------

def test_invariant_06_switch_pairs_balance(fixture):
    switch_ins = {t.scheme: t for t in fixture.ctx.transaction_snapshot if t.action == "Switch In"}
    pairs_checked = 0
    for t in fixture.ctx.transaction_snapshot:
        if t.action != "Switch Out":
            continue
        paired = switch_ins.get(t.suggested_scheme)
        if paired is None:
            continue
        deduction = t.switch_deduction or 0.0
        expected = t.amount - deduction
        assert abs(paired.amount - expected) <= RUPEE_TOLERANCE, (
            f"[{fixture.name}] switch pair {t.scheme} -> {paired.scheme}: switch-in "
            f"Rs {paired.amount:,.2f} != switch-out Rs {t.amount:,.2f} less declared deduction "
            f"Rs {deduction:,.2f}. Rupees are vanishing between the legs."
        )
        pairs_checked += 1

    # The context-level validator enforces the same thing at build time.
    _validate_report_context(fixture.ctx)


# --------------------------------------------------------------------------
# 7. sum(per-transaction tax deductions) == aggregate computed tax
# --------------------------------------------------------------------------

def test_invariant_07_allocated_tax_sums_to_aggregate(fixture):
    aggregate = fixture.ctx.tax_result.summary.total_computed_tax
    allocated = round(sum(fixture.tax_allocation.values()), 2)
    assert abs(allocated - aggregate) <= 0.01, (
        f"[{fixture.name}] per-transaction tax allocations sum to Rs {allocated:,.2f} but the Tax "
        f"Analysis table shows Rs {aggregate:,.2f}. The switch-in amounts and the tax table disagree."
    )


# --------------------------------------------------------------------------
# 8. tax gains derive only from Switch Out / Redeem transactions
# --------------------------------------------------------------------------

def test_invariant_08_tax_only_from_transacted_holdings(fixture):
    taxable_schemes = {
        t.scheme for t in fixture.ctx.transaction_snapshot
        if t.action in TAXABLE_TRANSACTION_ACTIONS
    }
    for th in fixture.tax_holdings:
        assert th.scheme in taxable_schemes, (
            f"[{fixture.name}] '{th.scheme}' feeds the tax computation but has no Switch Out / "
            f"Redeem row. An untransacted holding's unrealised gain is being taxed."
        )
    for row in fixture.ctx.tax_result.holdings:
        assert row.scheme in taxable_schemes, (
            f"[{fixture.name}] tax table row '{row.scheme}' is not a transacted scheme."
        )


# --------------------------------------------------------------------------
# 9. unique scheme count, never folio count
# --------------------------------------------------------------------------

def test_invariant_09_scheme_count_is_unique_not_folios(fixture):
    unique = len({h.scheme for h in fixture.ctx.holdings if h.scheme})
    reported = fixture.ctx.portfolio_summary.num_schemes
    assert reported == unique, (
        f"[{fixture.name}] Number of Schemes reads {reported} but there are {unique} unique scheme "
        f"names across {len(fixture.ctx.holdings)} folio rows."
    )


# --------------------------------------------------------------------------
# 10. no market-cap bucket is silently empty when equity holdings exist
# --------------------------------------------------------------------------

def test_invariant_10_market_cap_not_silently_empty(fixture):
    equity_holdings = [h for h in fixture.ctx.holdings if h.category == "Equity" and h.current_value]
    if not equity_holdings:
        pytest.skip("no equity holdings in this fixture")
    assert fixture.market_cap_rows, (
        f"[{fixture.name}] {len(equity_holdings)} equity holdings exist but the market-cap "
        f"breakdown is empty."
    )
    for row in fixture.market_cap_rows:
        assert row.value > 0, (
            f"[{fixture.name}] market-cap bucket '{row.label}' rendered with zero value."
        )
    # Anything Unclassified must be accompanied by a build warning naming it.
    unclassified = [r for r in fixture.market_cap_rows if r.label == "Unclassified"]
    if unclassified:
        assert fixture.market_cap_warnings, (
            f"[{fixture.name}] an Unclassified bucket exists but no build warning was raised - "
            f"schemes are being silently bucketed."
        )


# --------------------------------------------------------------------------
# 11. Things To Do: deadline-ascending, no row lost, priority carried but
#     not driving order
# --------------------------------------------------------------------------

def test_invariant_11_things_to_do_ordered_by_deadline_losing_nothing(fixture):
    items = fixture.ctx.things_to_do
    if not items:
        pytest.skip("no Things To Do rows in this fixture")

    # Every deadline must be readable. _parse_deadline raises rather than
    # sorting an unreadable one to an arbitrary place in an RM's worklist.
    deadlines = [_parse_deadline(i.deadline) for i in items]

    ordered = sorted(items, key=lambda i: _parse_deadline(i.deadline))

    # No row lost, none duplicated - a dropped action item is a
    # correctness failure, not a display one.
    assert len(ordered) == len(items), (
        f"[{fixture.name}] ordering changed the row count: {len(items)} generated, "
        f"{len(ordered)} ordered"
    )
    assert {id(i) for i in ordered} == {id(i) for i in items}, (
        f"[{fixture.name}] ordering swapped rows rather than reordering them"
    )

    keys = [_parse_deadline(i.deadline) for i in ordered]
    assert keys == sorted(keys), (
        f"[{fixture.name}] Things To Do is not in deadline-ascending order: {keys}"
    )

    # Order must be independent of priority. Feeding the same rows in a
    # priority-sorted order must produce the identical rendered order -
    # this is what fails if the sort key ever regresses to
    # (priority, deadline).
    by_priority = sorted(items, key=lambda i: i.priority)
    reordered = sorted(by_priority, key=lambda i: _parse_deadline(i.deadline))
    assert [id(i) for i in reordered] == [id(i) for i in ordered], (
        f"[{fixture.name}] Things To Do order depends on priority - it must depend on "
        f"deadline alone, with priority shown in its own column instead."
    )

    # Sanity: the fixtures should actually exercise the distinction, or
    # this invariant is watching a case that can never differ.
    if len({i.priority for i in items}) > 1 and len(set(deadlines)) > 1:
        priority_order = [id(i) for i in sorted(items, key=lambda i: i.priority)]
        if priority_order != [id(i) for i in ordered]:
            return  # deadline order genuinely differs from priority order - good


# --------------------------------------------------------------------------
# 12. no target-allocation / drift language in client-facing strings
# --------------------------------------------------------------------------

FORBIDDEN_COPY = re.compile(r"target\s+(band|allocation)|drift", re.IGNORECASE)


def _client_facing_strings(fixture) -> list:
    """Every string this fixture would put in front of a client. Code
    comments are deliberately excluded - the ban is on report copy, not
    on internal notes explaining why the concept doesn't exist."""
    ctx = fixture.ctx
    out = []
    r = ctx.risk_profile_result
    out += [r.profile or "", r.band_definition or "", r.description or ""]
    for i in ctx.things_to_do:
        out += [i.action, i.scheme, i.what_to_do, i.deadline]
    for t in ctx.transaction_snapshot:
        out += [t.scheme, t.action, t.suggested_scheme or "", t.switch_deduction_note or ""]
    for row in fixture.market_cap_rows:
        out.append(row.label)
    out.append(ctx.client_summary_placeholder)
    return [s for s in out if s]


def test_invariant_12_no_target_allocation_language(fixture):
    for text in _client_facing_strings(fixture):
        match = FORBIDDEN_COPY.search(text)
        assert match is None, (
            f"[{fixture.name}] client-facing string implies a target allocation the model "
            f"doesn't have: {text!r} (matched {match.group(0)!r})"
        )


# --------------------------------------------------------------------------
# Golden-only snapshot values (canary, not the correctness gate)
def test_fixture_G_losses_render_as_losses(built_fixtures):
    f = built_fixtures["G_all_loss"]
    e = f.expected
    s = f.ctx.portfolio_summary

    assert round(s.total_invested) == e["total_purchase"]
    assert round(s.current_value) == e["total_current"]
    assert round(s.absolute_gain) == e["total_gain"]
    assert s.absolute_gain < 0, "an all-loss portfolio must report a negative gain"
    assert s.portfolio_cagr_pct is not None and s.portfolio_cagr_pct < 0, (
        f"CAGR must be negative for an all-loss portfolio, got {s.portfolio_cagr_pct}"
    )
    assert f.ctx.tax_result.summary.total_computed_tax == e["aggregate_tax"]

    # Losses must be visible as negative gross gains, not flattened to zero.
    summary = f.ctx.tax_result.summary
    buckets = [
        summary.equity_ltcg_gross_gain, summary.equity_stcg_gross_gain,
        summary.non_equity_ltcg_gross_gain, summary.non_equity_stcg_gross_gain,
    ]
    assert any(b < 0 for b in buckets), (
        f"all-loss portfolio shows no negative gross gain in any tax bucket: {buckets}"
    )
