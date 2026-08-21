"""
pipeline/report_assembler.py

Turns one parsed ClientPortfolio into the ReportContext the builder
consumes. Pure wiring: every figure is produced by the same production
function the rest of the pipeline already uses, so nothing is computed
twice and nothing is computed here.

The one derivation this module does perform is purchase_date, and only
because HOLDING DAYS is a required column while a purchase date is in no
file at all - see _purchase_date_from_holding_days.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, timedelta
from pathlib import Path
from typing import Optional

from pipeline.dashboard_parser import ClientPortfolio, ParseWarning
from pipeline.docx_builder import (
    FirmInfo,
    PortfolioSummary,
    ReportContext,
    RMInfo,
    ThingsToDoRow,
    TransactionSnapshotRow,
)
from pipeline.parser import Holding, compute_value_weighted_cagr
from pipeline.risk_profile import (
    RiskHolding,
    classify_scheme_market_cap,
    compute_equity_market_cap_breakdown,
    compute_headline_equity_exposure_pct,
    compute_risk_profile,
)
from pipeline.tax_calc import (
    allocate_tax_across_transactions,
    build_tax_holdings_from_transactions,
    compute_portfolio_tax,
)

ASSETS_DIR = Path(__file__).parent.parent / "assets"

# The pie's buckets. Categories outside this map fall into Other/Liquid/Cash
# rather than creating a bucket of their own, so the pie can never show a
# slice the Holdings table has no rows for.
CATEGORY_TO_PIE_BUCKET = {
    "Equity": "Equity",
    "Hybrid": "Hybrid",
    "Debt": "Debt",
    "Gold": "Other/Liquid/Cash",
    "Other": "Other/Liquid/Cash",
    "Liquid": "Other/Liquid/Cash",
    "Arbitrage": "Other/Liquid/Cash",
}

# Days from the report date to each generated Things To Do deadline.
# Fixed offsets, so two runs of the same file produce the same worklist.
DEADLINE_OFFSET_DAYS = {
    "watchlist": 30,
    "accumulate": 30,
    "tax loss harvest": 45,
    "emergency_fund": 30,
    "kyc": 75,
}


def infer_fund_type(holding: Holding) -> str:
    """Free-text description that risk_profile's look-through WEIGHTING
    rules keyword-match on (how equity-like is this fund: 100% / 75% / 0%).

    Reads the SCHEME NAME first and the sheet's asset-class row only as a
    fallback. The name is the more specific signal - a Balanced Advantage
    Fund sitting under the "Hybrid" heading is 50% equity, not the 75% a
    bare "Hybrid" would imply, and the name is what says so.

    This is NOT market-cap classification; that is done separately from
    the scheme name by classify_scheme_market_cap().
    """
    name = (holding.scheme or "").lower()
    category = holding.category or ""

    for keyword in ("balanced advantage", "dynamic asset allocation", "multi asset",
                    "aggressive hybrid", "equity savings", "arbitrage", "liquid",
                    "overnight", "gilt", "gold", "silver", "index", "debt"):
        if keyword in name:
            return holding.scheme

    if category == "Equity":
        return holding.scheme or "Equity Fund"
    if category == "Hybrid":
        return "Aggressive Hybrid Fund"
    if category == "Debt":
        return "Debt Fund"
    if category == "Gold":
        return "Gold ETF"
    if category in ("Liquid", "Arbitrage"):
        return f"{category} Fund"
    if category == "Other":
        return "Gold ETF" if "gold" in name else "Other Fund"
    return holding.scheme or "Other Fund"


def _purchase_date_from_holding_days(holding: Holding, as_of: date) -> Optional[date]:
    """Derives a purchase date from HOLDING DAYS.

    No file on this data path carries a purchase date, but HOLDING DAYS is
    a required column - and holding period is the only thing the tax
    module needs the date FOR (LTCG vs STCG). Deriving it from days held
    uses data that is actually present instead of leaving every
    transaction unclassifiable. It is exact to the day the export was
    generated; if the file is stale by a week, a holding sitting within a
    week of the 12-month boundary could classify either way, which is why
    the tax section states its assumptions.
    """
    days = getattr(holding, "holding_days", None)
    if days is None:
        return None
    try:
        return as_of - timedelta(days=int(days))
    except (TypeError, ValueError, OverflowError):
        return None


@dataclass
class AssembledReport:
    ctx: ReportContext
    mindmap_recommendations: list = field(default_factory=list)
    market_cap_warnings: list = field(default_factory=list)
    assembly_warnings: list = field(default_factory=list)

    @property
    def unclassified_schemes(self) -> list:
        """Equity schemes the market-cap rules could not place. These become
        a per-scheme dropdown on the review screen rather than a console
        warning nobody reads."""
        return sorted({
            h.scheme for h in self.ctx.holdings
            if h.category == "Equity" and h.scheme
            and classify_scheme_market_cap(h.scheme) == "Unclassified"
        })


def build_transactions(client: ClientPortfolio, warnings: list) -> list:
    """Turns matched actions into Transaction Snapshot rows.

    Only the six permitted verbs are ever emitted. 'trim' becomes a
    Redeem, because a partial exit realises gains exactly as a full one
    does and the sheet carries no partial amount to size it with - which
    is itself warned about rather than assumed away.
    """
    rows: list = []
    by_key = {(h.scheme or "").lower(): h for h in client.holdings}

    for action in client.actions:
        if action.kind != "transaction":
            continue
        holding = by_key.get(action.scheme.lower())
        if holding is None:
            continue
        amount = holding.current_value

        if action.canonical == "switch":
            if not action.suggested_scheme:
                warnings.append(ParseWarning(
                    kind="switch_without_target",
                    message=(f"'{action.scheme}' is marked Switch but has no SUGGESTED SCHEME, so "
                             f"there is nowhere for the money to go. The row has been left out."),
                    sheet=action.sheet, row=action.row, client=client.name,
                ))
                continue
            rows.append(TransactionSnapshotRow(
                scheme=holding.scheme, action="Switch Out", amount=amount,
                balance_units=holding.balance_units, purchase_amount=holding.purchase_value,
                suggested_scheme=action.suggested_scheme,
                purchase_date=getattr(holding, "_purchase_date", None),
            ))
            rows.append(TransactionSnapshotRow(
                scheme=action.suggested_scheme, action="Switch In", amount=None,
            ))
        elif action.canonical in ("redeem", "trim"):
            if action.canonical == "trim":
                warnings.append(ParseWarning(
                    kind="trim_amount_assumed",
                    message=(f"'{action.scheme}' is marked Trim, which is a partial exit, but the "
                             f"sheet carries no amount. The full holding value has been used - "
                             f"confirm the intended amount before sending."),
                    sheet=action.sheet, row=action.row, client=client.name,
                ))
            rows.append(TransactionSnapshotRow(
                scheme=holding.scheme, action="Redeem", amount=amount,
                balance_units=holding.balance_units, purchase_amount=holding.purchase_value,
                purchase_date=getattr(holding, "_purchase_date", None),
            ))
    return rows


def build_things_to_do(client: ClientPortfolio, as_of: date) -> list:
    """Standing review items plus one row per non-transaction action."""
    def deadline(key: str) -> str:
        return (as_of + timedelta(days=DEADLINE_OFFSET_DAYS[key])).strftime("%d %b %Y")

    rows = [
        ThingsToDoRow(0, "Review", "Emergency Fund",
                      "Discuss liquid-fund buffer with client",
                      deadline("emergency_fund"), priority=1),
        ThingsToDoRow(0, "Update", "KYC / Nominee",
                      "Confirm nominee details are current across all folios",
                      deadline("kyc"), priority=4),
    ]
    label = {"watchlist": "Watchlist", "accumulate": "Accumulate",
             "tax loss harvest": "Tax Loss Harvest"}
    what = {
        "watchlist": "Keep under review - no action this cycle",
        "accumulate": "Add on dips / via SIP as discussed",
        "tax loss harvest": "Book the loss to set off against realised gains",
    }
    for action in client.actions:
        if action.kind != "things_to_do":
            continue
        rows.append(ThingsToDoRow(
            0, label[action.canonical], action.scheme, what[action.canonical],
            deadline(action.canonical), priority=2,
        ))
    return rows


def assemble_report_context(
    client: ClientPortfolio,
    *,
    as_of: date,
    rm: RMInfo,
    salutation: str = "Mr.",
    mindmap_path: Optional[Path] = None,
    market_cap_overrides: Optional[dict] = None,
) -> AssembledReport:
    """Builds the ReportContext for one client.

    market_cap_overrides: {scheme: category} decided by the RM on the
    review screen for schemes the rules could not classify. Supplied here
    rather than guessed, so an unclassified scheme is a question a human
    answered, not a default the pipeline picked.
    """
    from pipeline.mindmap import build_mindmap_recommendations_from_transactions

    warnings: list = []
    holdings = client.holdings

    for holding in holdings:
        holding._purchase_date = _purchase_date_from_holding_days(holding, as_of)

    asset_allocation: dict = {}
    for holding in holdings:
        bucket = CATEGORY_TO_PIE_BUCKET.get(holding.category, "Other/Liquid/Cash")
        asset_allocation[bucket] = asset_allocation.get(bucket, 0) + (holding.current_value or 0)

    risk_result = compute_risk_profile(
        [RiskHolding(h.scheme, infer_fund_type(h), h.current_value) for h in holdings]
    )
    market_cap_rows, market_cap_warnings = compute_equity_market_cap_breakdown(
        holdings, overrides=market_cap_overrides
    )

    transactions = build_transactions(client, warnings)

    tax_holdings, tax_warnings = build_tax_holdings_from_transactions(
        transactions, holdings, as_of=as_of
    )
    for message in tax_warnings:
        warnings.append(ParseWarning(kind="tax", message=message, client=client.name))

    tax_allocation = allocate_tax_across_transactions(tax_holdings, as_of=as_of)

    # Switch Out rows are paired with their Switch In POSITIONALLY, not by
    # target scheme name. build_transactions() emits each Switch In
    # directly after the Switch Out that funds it, so position is the
    # pairing - and it is the only thing that survives two switches into
    # the same target fund. Keying by scheme name collapsed both of
    # PRIYA SHARMA's Mirae switches onto one Switch In row: the first was
    # left with no amount, the second was written twice, and Rs 12.57 lakh
    # vanished from the report. _validate_report_context caught it at
    # build time, which is the only reason it was never sent.
    switch_in_for = {}
    pending_out = None
    for row in transactions:
        if row.action == "Switch Out":
            pending_out = row
        elif row.action == "Switch In" and pending_out is not None:
            switch_in_for[id(pending_out)] = row
            pending_out = None

    # A scheme's allocated tax covers every rupee leaving that scheme. When
    # the same fund is switched out of twice (two folios), the allocation
    # has to be SPLIT across those rows pro-rata by amount - charging each
    # row the scheme's full allocation would deduct the tax twice and
    # under-fund both switch-ins.
    out_rows_by_scheme: dict = {}
    for row in transactions:
        if row.action == "Switch Out":
            out_rows_by_scheme.setdefault(row.scheme, []).append(row)

    for scheme, rows_for_scheme in out_rows_by_scheme.items():
        allocated_total = tax_allocation.get(scheme, 0.0)
        gross = sum(r.amount for r in rows_for_scheme if r.amount is not None)
        for index, row in enumerate(rows_for_scheme):
            if len(rows_for_scheme) == 1 or not gross:
                share = allocated_total if index == 0 else 0.0
            elif index == len(rows_for_scheme) - 1:
                # Last row absorbs the rounding remainder so the split
                # sums to exactly the scheme's allocation.
                share = allocated_total - sum(
                    round(allocated_total * (r.amount or 0.0) / gross, 2)
                    for r in rows_for_scheme[:-1]
                )
            else:
                share = round(allocated_total * (row.amount or 0.0) / gross, 2)

            row.switch_deduction = share
            row.switch_deduction_note = "capital gains tax (this switch's share of the total)"
            paired = switch_in_for.get(id(row))
            if paired is not None and row.amount is not None:
                paired.amount = row.amount - share

    tax_result = compute_portfolio_tax(tax_holdings, as_of=as_of)

    total_invested = sum(h.purchase_value for h in holdings if h.purchase_value is not None)
    total_current = sum(h.current_value for h in holdings if h.current_value is not None)
    monthly_sip = sum(s.sip_amount for s in client.sips if s.sip_amount is not None)

    portfolio_summary = PortfolioSummary(
        total_invested=total_invested,
        current_value=total_current,
        absolute_gain=total_current - total_invested,
        absolute_gain_pct=((total_current - total_invested) / total_invested * 100)
        if total_invested else 0.0,
        portfolio_cagr_pct=compute_value_weighted_cagr(holdings),
        monthly_sip=monthly_sip,
        num_schemes=len({h.scheme for h in holdings if h.scheme}),
    )

    ctx = ReportContext(
        client_name=client.name.title(),
        client_salutation=salutation,
        report_date=as_of,
        firm=FirmInfo(),
        rm=rm,
        logo_path=ASSETS_DIR / "logo_2.jpg",
        portfolio_summary=portfolio_summary,
        asset_allocation=asset_allocation,
        equity_sub_allocation=market_cap_rows,
        risk_profile_result=risk_result,
        holdings=holdings,
        mindmap_path=mindmap_path or (ASSETS_DIR / "__no_mindmap__.png"),
        transaction_snapshot=transactions,
        performance_rows=[],
        tax_result=tax_result,
        emergency_fund_insurance=None,
        things_to_do=build_things_to_do(client, as_of),
        director_message_path=ASSETS_DIR / "director_message.docx",
        thank_you_message_path=ASSETS_DIR / "thank_you_message.docx",
    )

    return AssembledReport(
        ctx=ctx,
        mindmap_recommendations=build_mindmap_recommendations_from_transactions(transactions),
        market_cap_warnings=market_cap_warnings,
        assembly_warnings=warnings,
    )
