"""
pipeline/tax_calc.py

Capital-gains tax computation for the WC Securities / Wealthkare Portfolio
Review pipeline, per Tax Reckoner 2026-27.

Scope
-----
- LTCG / STCG classification per holding, from purchase date + holding
  period as of a valuation/sale date (equity-oriented: >12 months = LTCG;
  non-equity: >24 months = LTCG).
- Tax Reckoner 2026-27 rates:
    Equity LTCG (s.112A):  12.5% on gains above Rs 1.25 lakh (aggregate,
                            per financial year)
    Equity STCG:            20%
    Non-equity LTCG:        12.5% (no indexation)
    Non-equity STCG:        applicable slab rates (not computable here -
                             depends on the client's total taxable income,
                             which this module does not have)
    Health & Education Cess: 4% on the base tax, on all of the above.
- The Rs 1.25 lakh equity LTCG exemption is applied ONCE, in aggregate,
  across all equity-LTCG gains passed into a single compute_portfolio_tax()
  call - not re-applied per holding. Every LTCG holding still carries the
  standard assumption-flag text (see ASSUMPTION_FLAG_TEXT below), since a
  single report may not see redemptions the client makes elsewhere in the
  same FY.
- Advance-tax reminder: any gain whose sale/booking date falls in
  July-September (Q2 of the FY) is flagged.
- Tax-loss harvesting candidates: equity-oriented or non-equity holdings
  purchased within the last 12 months that are currently sitting at an
  unrealised loss.
  caller. Anything not supplied is flagged "not available in uploaded
  file" rather than assumed to be zero or maxed out.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from typing import Optional

# --------------------------------------------------------------------------
# Constants (Tax Reckoner 2026-27)
# --------------------------------------------------------------------------

EQUITY_LTCG_HOLDING_MONTHS = 12
NON_EQUITY_LTCG_HOLDING_MONTHS = 24

EQUITY_LTCG_EXEMPTION = 125_000.0       # Rs 1.25 lakh, s.112A, aggregate per FY
EQUITY_LTCG_RATE = 0.125                # 12.5% above the exemption
EQUITY_STCG_RATE = 0.20                 # 20%
NON_EQUITY_LTCG_RATE = 0.125            # 12.5%, no indexation

CESS_RATE = 0.04                        # Health & Education Cess, on base tax

ASSUMPTION_FLAG_TEXT = (
    "The above LTCG estimate assumes the annual ₹1.25 lakh exemption is "
    "otherwise unused in this financial year. If multiple fund redemptions "
    "occur in the same FY, the exemption may be exhausted — consult "
    "before executing all transactions."
)

ADVANCE_TAX_FLAG_TEMPLATE = "Advance tax due by 15 September {year}."

Q2_MONTHS = (7, 8, 9)  # July, August, September



# --------------------------------------------------------------------------
# Data model
# --------------------------------------------------------------------------

@dataclass
class TaxHolding:
    scheme: str
    is_equity_oriented: bool
    purchase_date: date
    purchase_value: float
    current_value: float  # treated as the redemption/valuation value
    sale_date: Optional[date] = None  # defaults to as_of date if not given
    member: Optional[str] = None


@dataclass
class HoldingTaxResult:
    scheme: str
    member: Optional[str]
    is_equity_oriented: bool
    holding_period_months: int
    classification: str  # "LTCG" or "STCG"
    gain: float           # positive = gain, negative = loss
    is_loss: bool
    assumption_flag: Optional[str] = None
    advance_tax_flag: Optional[str] = None
    warnings: list[str] = field(default_factory=list)


@dataclass
class TaxSummary:
    equity_ltcg_gross_gain: float = 0.0
    equity_ltcg_exemption_applied: float = 0.0
    equity_ltcg_taxable_gain: float = 0.0
    equity_ltcg_base_tax: float = 0.0
    equity_ltcg_cess: float = 0.0
    equity_ltcg_total_tax: float = 0.0

    equity_stcg_gross_gain: float = 0.0
    equity_stcg_base_tax: float = 0.0
    equity_stcg_cess: float = 0.0
    equity_stcg_total_tax: float = 0.0

    non_equity_ltcg_gross_gain: float = 0.0
    non_equity_ltcg_base_tax: float = 0.0
    non_equity_ltcg_cess: float = 0.0
    non_equity_ltcg_total_tax: float = 0.0

    non_equity_stcg_gross_gain: float = 0.0
    non_equity_stcg_note: str = "Slab rate applies — not computed here (depends on total taxable income)."

    total_computed_tax: float = 0.0


@dataclass
class PortfolioTaxResult:
    holdings: list[HoldingTaxResult] = field(default_factory=list)
    summary: TaxSummary = field(default_factory=TaxSummary)
    tlh_opportunities: list[HoldingTaxResult] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


# --------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------

def _months_between(start: date, end: date) -> int:
    """Whole calendar months between two dates (end - start), matching how
    holding periods are counted for LTCG/STCG purposes."""
    months = (end.year - start.year) * 12 + (end.month - start.month)
    if end.day < start.day:
        months -= 1
    return max(months, 0)


def _round2(value: float) -> float:
    return round(value, 2)


# --------------------------------------------------------------------------
# Per-holding classification
# --------------------------------------------------------------------------

def classify_holding(h: TaxHolding, as_of: date) -> HoldingTaxResult:
    warnings: list[str] = []

    if h.purchase_date is None:
        warnings.append(f"'{h.scheme}': purchase date missing — cannot classify LTCG/STCG.")
        return HoldingTaxResult(
            scheme=h.scheme, member=h.member, is_equity_oriented=h.is_equity_oriented,
            holding_period_months=0, classification="UNKNOWN", gain=0.0, is_loss=False,
            warnings=warnings,
        )

    sale_date = h.sale_date or as_of
    if sale_date < h.purchase_date:
        warnings.append(
            f"'{h.scheme}': sale/valuation date ({sale_date}) is before purchase date "
            f"({h.purchase_date}) — skipping classification rather than guessing."
        )
        return HoldingTaxResult(
            scheme=h.scheme, member=h.member, is_equity_oriented=h.is_equity_oriented,
            holding_period_months=0, classification="UNKNOWN", gain=0.0, is_loss=False,
            warnings=warnings,
        )

    holding_months = _months_between(h.purchase_date, sale_date)
    threshold = EQUITY_LTCG_HOLDING_MONTHS if h.is_equity_oriented else NON_EQUITY_LTCG_HOLDING_MONTHS
    classification = "LTCG" if holding_months > threshold else "STCG"

    if h.purchase_value is None or h.current_value is None:
        warnings.append(f"'{h.scheme}': purchase or current value missing — gain not computed.")
        gain = 0.0
    else:
        gain = h.current_value - h.purchase_value

    is_loss = gain < 0

    assumption_flag = ASSUMPTION_FLAG_TEXT if classification == "LTCG" else None

    advance_tax_flag = None
    if not is_loss and gain > 0 and sale_date.month in Q2_MONTHS:
        advance_tax_flag = ADVANCE_TAX_FLAG_TEMPLATE.format(year=sale_date.year)

    return HoldingTaxResult(
        scheme=h.scheme,
        member=h.member,
        is_equity_oriented=h.is_equity_oriented,
        holding_period_months=holding_months,
        classification=classification,
        gain=_round2(gain),
        is_loss=is_loss,
        assumption_flag=assumption_flag,
        advance_tax_flag=advance_tax_flag,
        warnings=warnings,
    )


# --------------------------------------------------------------------------
# Portfolio-level aggregation
# --------------------------------------------------------------------------

def _aggregate_tax(results: list[HoldingTaxResult]) -> TaxSummary:
    summary = TaxSummary()

    for r in results:
        if r.classification not in ("LTCG", "STCG"):
            continue  # UNKNOWN rows (e.g. missing purchase date) don't contribute

        # Losses ARE included here (r.gain can be negative) so they NET
        # against gains within the same bucket, rather than being dropped.
        # A proposed transaction that's a loss (e.g. a short-term Redeem at
        # a loss) must still show up - as a loss available for set-off,
        # visible in its bucket's gross gain figure - not vanish silently.
        if r.is_equity_oriented and r.classification == "LTCG":
            summary.equity_ltcg_gross_gain += r.gain
        elif r.is_equity_oriented and r.classification == "STCG":
            summary.equity_stcg_gross_gain += r.gain
        elif not r.is_equity_oriented and r.classification == "LTCG":
            summary.non_equity_ltcg_gross_gain += r.gain
        elif not r.is_equity_oriented and r.classification == "STCG":
            summary.non_equity_stcg_gross_gain += r.gain

    # Equity LTCG: Rs 1.25L exemption applied once, in aggregate. A net
    # LOSS in this bucket (gross_gain < 0) consumes no exemption and owes
    # no tax - max(0, ...) handles both that and the ordinary gain case.
    exemption_applied = min(max(0.0, summary.equity_ltcg_gross_gain), EQUITY_LTCG_EXEMPTION)
    taxable = max(0.0, summary.equity_ltcg_gross_gain - EQUITY_LTCG_EXEMPTION)
    summary.equity_ltcg_exemption_applied = _round2(exemption_applied)
    summary.equity_ltcg_taxable_gain = _round2(taxable)
    summary.equity_ltcg_base_tax = _round2(taxable * EQUITY_LTCG_RATE)
    summary.equity_ltcg_cess = _round2(summary.equity_ltcg_base_tax * CESS_RATE)
    summary.equity_ltcg_total_tax = _round2(summary.equity_ltcg_base_tax + summary.equity_ltcg_cess)

    # Equity STCG: flat 20%, no exemption. A net loss owes zero tax (the
    # loss itself is still visible via gross_gain, which is allowed to be
    # negative) - max(0, ...) on the TAXABLE base only, not the displayed
    # gross figure.
    equity_stcg_taxable = max(0.0, summary.equity_stcg_gross_gain)
    summary.equity_stcg_base_tax = _round2(equity_stcg_taxable * EQUITY_STCG_RATE)
    summary.equity_stcg_cess = _round2(summary.equity_stcg_base_tax * CESS_RATE)
    summary.equity_stcg_total_tax = _round2(summary.equity_stcg_base_tax + summary.equity_stcg_cess)

    # Non-equity LTCG: flat 12.5%, no indexation, no exemption threshold specified.
    non_equity_ltcg_taxable = max(0.0, summary.non_equity_ltcg_gross_gain)
    summary.non_equity_ltcg_base_tax = _round2(non_equity_ltcg_taxable * NON_EQUITY_LTCG_RATE)
    summary.non_equity_ltcg_cess = _round2(summary.non_equity_ltcg_base_tax * CESS_RATE)
    summary.non_equity_ltcg_total_tax = _round2(summary.non_equity_ltcg_base_tax + summary.non_equity_ltcg_cess)

    # Non-equity STCG: slab rate - left uncomputed, gross gain still reported.
    summary.non_equity_stcg_gross_gain = _round2(summary.non_equity_stcg_gross_gain)

    summary.equity_ltcg_gross_gain = _round2(summary.equity_ltcg_gross_gain)
    summary.equity_stcg_gross_gain = _round2(summary.equity_stcg_gross_gain)
    summary.non_equity_ltcg_gross_gain = _round2(summary.non_equity_ltcg_gross_gain)

    summary.total_computed_tax = _round2(
        summary.equity_ltcg_total_tax
        + summary.equity_stcg_total_tax
        + summary.non_equity_ltcg_total_tax
        # non-equity STCG intentionally excluded - slab-dependent, not computed
    )

    return summary


def _find_tlh_opportunities(
    holdings: list[TaxHolding], results: list[HoldingTaxResult], as_of: date
) -> list[HoldingTaxResult]:
    """Unrealised SHORT-TERM losses: purchased within 12 months of `as_of`
    and currently at a loss. (Long-term losses are also harvestable in
    practice, but this module reports short-term candidates specifically,
    per spec, since those are the most time-sensitive to act on.)"""
    opportunities = []
    for h, r in zip(holdings, results):
        if r.classification == "UNKNOWN":
            continue
        if not r.is_loss:
            continue
        months_held = _months_between(h.purchase_date, h.sale_date or as_of)
        if months_held <= 12:
            opportunities.append(r)
    return opportunities


def compute_portfolio_tax(
    holdings: list[TaxHolding],
    as_of: Optional[date] = None,
) -> PortfolioTaxResult:
    as_of = as_of or date.today()

    result_rows = [classify_holding(h, as_of) for h in holdings]

    warnings: list[str] = []
    for r in result_rows:
        warnings.extend(r.warnings)

    summary = _aggregate_tax(result_rows)
    tlh_opportunities = _find_tlh_opportunities(holdings, result_rows, as_of)

    return PortfolioTaxResult(
        holdings=result_rows,
        summary=summary,
        tlh_opportunities=tlh_opportunities,
        warnings=warnings,
    )


# --------------------------------------------------------------------------
# Building tax holdings FROM the Transaction Snapshot (not a separate list)
# --------------------------------------------------------------------------

TAXABLE_TRANSACTION_ACTIONS = {"Switch Out", "Redeem"}


def build_tax_holdings_from_transactions(transactions, holdings, as_of: date) -> tuple:
    """Builds the TaxHolding list used for capital-gains computation SOLELY
    from the proposed Switch Out / Redeem rows of the Transaction Snapshot -
    per the rule that this section taxes only what's actually being
    transacted. Unrealised gains on holdings nobody is touching (no
    Switch Out / Redeem row) are never taxed here.

    transactions: the SAME list rendered in the Transaction Snapshot table
        (docx_builder.TransactionSnapshotRow instances, duck-typed here to
        avoid a circular import - each needs .scheme, .action, .amount
        (CURRENT value of the units transacted), .purchase_amount, and
        .purchase_date).
    holdings: pipeline.parser.Holding objects, used only to look up each
        transacted scheme's category (Equity/Hybrid -> equity-oriented for
        tax purposes; Debt/Other -> not).
    as_of: used as the sale/valuation date for every transaction (they're
        all proposed to execute now).

    Returns (tax_holdings, warnings). A transaction is skipped - with a
    warning, never guessed - if it has no purchase_date (required for
    LTCG/STCG classification) or no purchase/current amount.
    """
    holdings_by_scheme = {h.scheme: h for h in holdings}
    tax_holdings: list[TaxHolding] = []
    warnings: list[str] = []

    for t in transactions:
        if t.action not in TAXABLE_TRANSACTION_ACTIONS:
            continue

        purchase_date = getattr(t, "purchase_date", None)
        if purchase_date is None:
            warnings.append(
                f"'{t.scheme}' ({t.action}): no purchase date on this transaction - "
                f"excluded from the tax computation rather than guessed."
            )
            continue

        purchase_value = t.purchase_amount
        current_value = t.amount  # Amount = current value of the units being transacted
        if purchase_value is None or current_value is None:
            warnings.append(
                f"'{t.scheme}' ({t.action}): missing Purchase Amount or Amount - "
                f"excluded from the tax computation."
            )
            continue

        holding = holdings_by_scheme.get(t.scheme)
        is_equity_oriented = holding.category in ("Equity", "Hybrid") if holding else True

        tax_holdings.append(TaxHolding(
            scheme=t.scheme,
            is_equity_oriented=is_equity_oriented,
            purchase_date=purchase_date,
            purchase_value=purchase_value,
            current_value=current_value,
            sale_date=as_of,
        ))

    # Regression guard for the "gains feeding the tax table are derived
    # from the same transaction list" requirement. True by construction
    # (the loop above only ever reads from `transactions`), asserted
    # explicitly so a future refactor that breaks the invariant fails
    # loudly at build time instead of silently drifting.
    taxable_scheme_names = {t.scheme for t in transactions if t.action in TAXABLE_TRANSACTION_ACTIONS}
    built_scheme_names = {th.scheme for th in tax_holdings}
    assert built_scheme_names.issubset(taxable_scheme_names), (
        "build_tax_holdings_from_transactions produced a TaxHolding for a scheme not present "
        "in the Transaction Snapshot's Switch Out/Redeem rows - this should be impossible."
    )

    return tax_holdings, warnings


# --------------------------------------------------------------------------
# Apportioning the AGGREGATE tax bill across individual transactions
# --------------------------------------------------------------------------

def allocate_tax_across_transactions(tax_holdings: list, as_of: date) -> dict:
    """Computes the total capital-gains tax across ALL proposed
    transactions once, then apportions it back to each transaction
    pro-rata by that transaction's positive gain.

    This replaces an earlier per-transaction estimator that applied the
    Rs 1.25 lakh s.112A exemption independently to every transaction. That
    exemption is ANNUAL and SHARED across the whole financial year, so
    per-transaction application understates tax whenever more than one
    transaction is proposed: two equity-LTCG switches with a Rs 80,000
    gain each would each individually look exempt (Rs 0 tax), while the
    aggregate computation correctly taxes Rs 35,000 of combined gain. The
    switch-in amounts and the Tax Analysis table would then disagree.

    Returns {scheme: allocated_tax}, where the allocated amounts sum to
    exactly PortfolioTaxResult.summary.total_computed_tax (any rounding
    remainder is pushed onto the largest-gain transaction so the total
    always reconciles to the rupee).

    Note this deliberately apportions across every taxable transaction,
    not just switches - a Redeem's gain consumes shared exemption too, so
    excluding it would over-allocate tax to the switches.
    """
    results = [classify_holding(h, as_of) for h in tax_holdings]
    summary = _aggregate_tax(results)
    total_tax = summary.total_computed_tax

    gains = {}
    for h, r in zip(tax_holdings, results):
        if r.classification in ("LTCG", "STCG") and r.gain > 0:
            gains[h.scheme] = gains.get(h.scheme, 0.0) + r.gain

    total_gain = sum(gains.values())
    if total_tax <= 0 or total_gain <= 0:
        return {scheme: 0.0 for scheme in gains}

    allocations = {
        scheme: _round2(total_tax * (gain / total_gain))
        for scheme, gain in gains.items()
    }

    # Push any rounding remainder onto the largest-gain transaction so the
    # allocations reconcile exactly to the aggregate figure shown in the
    # Tax Analysis table.
    remainder = _round2(total_tax - sum(allocations.values()))
    if remainder and allocations:
        largest = max(gains, key=lambda s: gains[s])
        allocations[largest] = _round2(allocations[largest] + remainder)

    return allocations


# --------------------------------------------------------------------------
# Self-test (dummy holdings, no external files required)
# --------------------------------------------------------------------------

def _run_self_test() -> None:
    print("=== pipeline/tax_calc.py self-test ===\n")

    as_of = date(2026, 8, 16)  # "today" for this test run

    holdings = [
        # Equity LTCG, large gain -> should push past the 1.25L exemption
        TaxHolding(
            scheme="Axis Bluechip Fund", is_equity_oriented=True,
            purchase_date=date(2023, 1, 10), purchase_value=500_000, current_value=750_000,
            sale_date=date(2026, 8, 1),
        ),
        # Equity LTCG, smaller gain, sold in August 2026 -> also equity LTCG bucket
        TaxHolding(
            scheme="Mirae Asset Large Cap Fund", is_equity_oriented=True,
            purchase_date=date(2024, 3, 15), purchase_value=200_000, current_value=260_000,
            sale_date=date(2026, 8, 5),
        ),
        # Equity STCG, sold in Q2 (Aug) -> advance tax flag expected
        TaxHolding(
            scheme="Quant Small Cap Fund", is_equity_oriented=True,
            purchase_date=date(2026, 3, 1), purchase_value=100_000, current_value=118_000,
            sale_date=date(2026, 8, 10),
        ),
        # Equity short-term LOSS, purchased <12 months ago -> TLH candidate
        TaxHolding(
            scheme="Nippon India Small Cap Fund", is_equity_oriented=True,
            purchase_date=date(2026, 2, 1), purchase_value=90_000, current_value=72_000,
            sale_date=None,  # not yet sold - valued as of `as_of`
        ),
        # Non-equity LTCG (debt fund held > 24 months)
        TaxHolding(
            scheme="ICICI Pru Short Term Debt Fund", is_equity_oriented=False,
            purchase_date=date(2022, 5, 1), purchase_value=300_000, current_value=345_000,
            sale_date=date(2026, 6, 15),
        ),
        # Non-equity STCG (debt fund held < 24 months) -> slab rate, not computed
        TaxHolding(
            scheme="HDFC Corporate Bond Fund", is_equity_oriented=False,
            purchase_date=date(2025, 6, 1), purchase_value=150_000, current_value=157_000,
            sale_date=date(2026, 7, 20),
        ),
        # Missing purchase date -> should warn, not crash
        TaxHolding(
            scheme="Unknown Legacy Folio", is_equity_oriented=True,
            purchase_date=None, purchase_value=50_000, current_value=55_000,
        ),
    ]


    result = compute_portfolio_tax(
        holdings, as_of=as_of,
    )

    print("--- Per-holding classification ---")
    for r in result.holdings:
        print(
            f"  {r.scheme:<32} {r.classification:<8} months={r.holding_period_months:<4} "
            f"gain={r.gain:>12,.2f}  loss={r.is_loss}"
        )
        if r.assumption_flag:
            print(f"      [assumption] {r.assumption_flag}")
        if r.advance_tax_flag:
            print(f"      [advance tax] {r.advance_tax_flag}")
        for w in r.warnings:
            print(f"      [warning] {w}")

    print("\n--- Portfolio tax summary ---")
    s = result.summary
    print(f"Equity LTCG gross gain:      {s.equity_ltcg_gross_gain:>12,.2f}")
    print(f"  Exemption applied:         {s.equity_ltcg_exemption_applied:>12,.2f}")
    print(f"  Taxable gain:              {s.equity_ltcg_taxable_gain:>12,.2f}")
    print(f"  Base tax (12.5%):          {s.equity_ltcg_base_tax:>12,.2f}")
    print(f"  Cess (4%):                 {s.equity_ltcg_cess:>12,.2f}")
    print(f"  Total tax:                 {s.equity_ltcg_total_tax:>12,.2f}")
    print(f"Equity STCG gross gain:      {s.equity_stcg_gross_gain:>12,.2f}")
    print(f"  Total tax (20% + cess):    {s.equity_stcg_total_tax:>12,.2f}")
    print(f"Non-equity LTCG gross gain:  {s.non_equity_ltcg_gross_gain:>12,.2f}")
    print(f"  Total tax (12.5% + cess):  {s.non_equity_ltcg_total_tax:>12,.2f}")
    print(f"Non-equity STCG gross gain:  {s.non_equity_stcg_gross_gain:>12,.2f}  ({s.non_equity_stcg_note})")
    print(f"TOTAL COMPUTED TAX:          {s.total_computed_tax:>12,.2f}")

    print(f"\n--- Tax-loss harvesting opportunities ({len(result.tlh_opportunities)}) ---")
    for r in result.tlh_opportunities:
        print(f"  {r.scheme}: unrealised loss of {abs(r.gain):,.2f} (held {r.holding_period_months} months)")

    print("\n--- allocate_tax_across_transactions ---")

    # Two equity-LTCG switches, Rs 80,000 gain each. Individually each
    # looks exempt (80,000 < 1.25L). Together they realise Rs 1,60,000, of
    # which Rs 35,000 is taxable - so real tax IS payable and both
    # switches must carry a share of it.
    two_switches = [
        TaxHolding("Switch A", True, date(2024, 1, 15), 200_000, 280_000, sale_date=as_of),
        TaxHolding("Switch B", True, date(2024, 2, 20), 300_000, 380_000, sale_date=as_of),
    ]
    allocations = allocate_tax_across_transactions(two_switches, as_of=as_of)
    aggregate = compute_portfolio_tax(two_switches, as_of=as_of).summary.total_computed_tax

    expected_aggregate = round(max(0.0, 160_000 - 125_000) * 0.125 * 1.04, 2)
    print(f"Combined gain 1,60,000 -> aggregate tax {aggregate} (expected {expected_aggregate})")
    print(f"Per-switch allocations: {allocations}")
    assert aggregate == expected_aggregate, f"Expected {expected_aggregate}, got {aggregate}"
    assert aggregate > 0, "Two switches over the shared exemption must produce real tax."

    allocated_total = round(sum(allocations.values()), 2)
    assert allocated_total == aggregate, (
        f"Sum of per-transaction deductions ({allocated_total}) must equal the aggregate "
        f"tax ({aggregate}) - otherwise the switch-in amounts and the Tax Analysis table disagree."
    )
    # Equal gains -> equal split
    assert allocations["Switch A"] == allocations["Switch B"], (
        "Two transactions with identical gains should carry identical shares of the tax."
    )
    print(f"Sum of allocations {allocated_total} == aggregate tax {aggregate}")

    # A single small switch under the exemption still owes nothing.
    one_small_switch = [TaxHolding("Small", True, date(2024, 1, 15), 250_000, 298_000, sale_date=as_of)]
    small_alloc = allocate_tax_across_transactions(one_small_switch, as_of=as_of)
    print(f"Single 48,000-gain switch (under exemption): {small_alloc}")
    assert sum(small_alloc.values()) == 0.0

    # A loss-making transaction is never allocated tax.
    with_loss = [
        TaxHolding("Winner", True, date(2024, 1, 15), 200_000, 480_000, sale_date=as_of),
        TaxHolding("Loser", True, date(2026, 3, 1), 80_000, 71_500, sale_date=as_of),
    ]
    loss_alloc = allocate_tax_across_transactions(with_loss, as_of=as_of)
    print(f"Winner/Loser allocations: {loss_alloc}")
    assert loss_alloc.get("Loser", 0.0) == 0.0, "A loss-making transaction must never be allocated tax."

    print("\nAll self-test assertions passed.")


if __name__ == "__main__":
    _run_self_test()
