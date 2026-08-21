"""
tests/fixtures.py

Standalone input datasets for the regression suite, each with
hand-computed expected values committed alongside the data.

Design rule
-----------
`make_context()` derives every downstream value by calling the SAME
production functions the real pipeline calls - compute_risk_profile(),
compute_headline_equity_exposure_pct(), compute_equity_market_cap_
breakdown(), build_tax_holdings_from_transactions(),
allocate_tax_across_transactions(), compute_portfolio_tax(),
compute_value_weighted_cagr(). It never reimplements them.

That distinction matters: if this helper computed, say, asset_allocation
its own way, the invariant tests would be checking the test helper
against itself and would pass while production was broken. Everything
here is either raw input data or a call into pipeline/.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
from typing import Optional

from pipeline.docx_builder import (
    FirmInfo, PerformanceRow, PortfolioSummary, RMInfo, ReportContext,
    ThingsToDoRow, TransactionSnapshotRow,
)
from pipeline.parser import Holding, compute_value_weighted_cagr
from pipeline.risk_profile import (
    RiskHolding, classify_scheme_market_cap, compute_equity_market_cap_breakdown,
    compute_headline_equity_exposure_pct, compute_risk_profile,
)
from pipeline.tax_calc import (
    allocate_tax_across_transactions, build_tax_holdings_from_transactions,
    compute_portfolio_tax,
)

ASSETS_DIR = Path(__file__).parent.parent / "assets"
AS_OF = date(2026, 8, 16)

# Pie buckets, keyed by the Holding.category values the parser emits.
CATEGORY_TO_PIE_BUCKET = {
    "Equity": "Equity", "Hybrid": "Hybrid", "Debt": "Debt", "Other": "Gold/SGB",
}


@dataclass
class Fixture:
    """One dataset plus everything the invariants need to check it."""
    name: str
    ctx: ReportContext
    mindmap_recs: list
    tax_allocation: dict          # {scheme: allocated tax}
    tax_holdings: list            # TaxHolding objects fed to the tax table
    market_cap_rows: list
    market_cap_warnings: list
    equity_exposure_pct: Optional[float]
    expected: dict = field(default_factory=dict)   # hand-computed, committed


def _mk_holding(category, scheme, folio, units, purchase, current, cagr=None) -> Holding:
    gain = current - purchase
    return Holding(
        member="Test Client", pan="ABCDE1234F", category=category,
        scheme=scheme, folio=folio, balance_units=units,
        purchase_value=purchase, current_value=current, gain=gain,
        absolute_return_pct=round(gain / purchase * 100, 2) if purchase else None,
        cagr_pct=cagr if cagr is not None else (round(gain / purchase * 100 / 3, 2) if purchase else None),
    )


def _infer_fund_type(h: Holding) -> str:
    """Maps a Holding to the free-text fund_type that risk_profile's
    look-through WEIGHTING rules key off (100% / 75% / 0% equity). This is
    NOT market-cap classification - that's done from the scheme name by
    classify_scheme_market_cap()."""
    if h.category == "Equity":
        return "Large Cap Fund"
    if h.category == "Hybrid":
        return "Aggressive Hybrid Fund"
    if h.category == "Debt":
        return "Debt Fund"
    if h.category == "Other":
        return "Gold ETF" if "gold" in (h.scheme or "").lower() else "Other Fund"
    return "Other Fund"


def make_context(
    name: str,
    holdings: list,
    transactions: list,
    *,
    performance_rows: Optional[list] = None,
    things_to_do: Optional[list] = None,
    expected: Optional[dict] = None,
    mindmap_path: Optional[Path] = None,
) -> Fixture:
    """Builds a full ReportContext from raw holdings + transactions, using
    production derivation functions throughout."""
    from pipeline.chart_gen import format_inr
    from pipeline.mindmap import build_mindmap_recommendations_from_transactions

    # --- asset allocation: pure aggregation of holdings, no extra buckets ---
    asset_allocation: dict = {}
    for h in holdings:
        bucket = CATEGORY_TO_PIE_BUCKET.get(h.category, "Other/Liquid/Cash")
        asset_allocation[bucket] = asset_allocation.get(bucket, 0) + (h.current_value or 0)

    risk_result = compute_risk_profile(
        [RiskHolding(h.scheme, _infer_fund_type(h), h.current_value) for h in holdings]
    )
    equity_exposure_pct = compute_headline_equity_exposure_pct(holdings)
    market_cap_rows, market_cap_warnings = compute_equity_market_cap_breakdown(holdings)

    # --- tax: derived from the transaction list, allocated pro-rata ---
    tax_holdings, _ = build_tax_holdings_from_transactions(transactions, holdings, as_of=AS_OF)
    tax_allocation = allocate_tax_across_transactions(tax_holdings, as_of=AS_OF)

    # Fill Switch In amounts = paired Switch Out less its allocated tax.
    switch_in_by_scheme = {t.scheme: t for t in transactions if t.action == "Switch In"}
    for t in transactions:
        if t.action != "Switch Out":
            continue
        allocated = tax_allocation.get(t.scheme, 0.0)
        t.switch_deduction = allocated
        t.switch_deduction_note = "capital gains tax (this switch's share of the total)"
        paired = switch_in_by_scheme.get(t.suggested_scheme)
        if paired is not None:
            paired.amount = t.amount - allocated

    tax_result = compute_portfolio_tax(tax_holdings, as_of=AS_OF)

    total_invested = sum(h.purchase_value for h in holdings if h.purchase_value is not None)
    total_current = sum(h.current_value for h in holdings if h.current_value is not None)
    portfolio_summary = PortfolioSummary(
        total_invested=total_invested,
        current_value=total_current,
        absolute_gain=total_current - total_invested,
        absolute_gain_pct=(total_current - total_invested) / total_invested * 100 if total_invested else 0.0,
        portfolio_cagr_pct=compute_value_weighted_cagr(holdings),
        monthly_sip=25000,
        num_schemes=len({h.scheme for h in holdings if h.scheme}),
    )

    if things_to_do is None:
        things_to_do = [
            ThingsToDoRow(0, "Review", "Emergency Fund",
                          "Discuss liquid-fund buffer with client", "15 Sep 2026", priority=1),
            ThingsToDoRow(0, "Update", "KYC / Nominee",
                          "Confirm nominee details are current", "31 Oct 2026", priority=4),
        ]

    ctx = ReportContext(
        client_name="Test Client",
        client_salutation="Mr.",
        report_date=AS_OF,
        firm=FirmInfo(),
        rm=RMInfo(name="Relationship Manager", email="rm@wealthcareindia.com", phone="+91-98100-00000"),
        logo_path=ASSETS_DIR / "logo_2.jpg",
        portfolio_summary=portfolio_summary,
        asset_allocation=asset_allocation,
        equity_sub_allocation=market_cap_rows,
        risk_profile_result=risk_result,
        holdings=holdings,
        mindmap_path=mindmap_path or (ASSETS_DIR / "__no_such_mindmap__.png"),
        transaction_snapshot=transactions,
        performance_rows=performance_rows or [],
        tax_result=tax_result,
        emergency_fund_insurance=None,
        things_to_do=things_to_do,
        director_message_path=ASSETS_DIR / "director_message.docx",
        thank_you_message_path=ASSETS_DIR / "thank_you_message.docx",
    )

    return Fixture(
        name=name, ctx=ctx,
        mindmap_recs=build_mindmap_recommendations_from_transactions(transactions),
        tax_allocation=tax_allocation,
        tax_holdings=tax_holdings,
        market_cap_rows=market_cap_rows,
        market_cap_warnings=market_cap_warnings,
        equity_exposure_pct=equity_exposure_pct,
        expected=expected or {},
    )


# ==========================================================================
# A. golden - the current Rahul Sharma dataset
# ==========================================================================

_GOLDEN_EQUITY = [
    ("Axis Bluechip Fund", "477288232357", 1245.678, 500000, 612000),
    ("Axis Focused Fund Reg (G)", "910855361015", 3200.442, 350000, 402340),
    ("HDFC Flexicap Fund", "118823409981", 980.234, 250000, 298000),
    ("HDFC Large And Mid Cap Fund Reg (G)", "229439832712", 4210.887, 420000, 486210),
    ("HDFC Mid Cap Fund Reg (G)", "152452455634", 1890.221, 300000, 398760),
    ("ICICI Pru Large Cap Fund Reg (G)", "103909552587", 5120.909, 600000, 712450),
    ("Invesco India Contra Fund Reg (G)", "310115421369", 2870.335, 350000, 421980),
    ("Kotak Flexi Cap Fund Reg (G)", "577996412", 1980.556, 280000, 334120),
    ("Kotak Large & Midcap Fund Reg (G)", "72627351", 2210.774, 310000, 379540),
    ("L&T Emerging Businesses Fund", "990012345678", 560.912, 80000, 71500),
    ("Canara Robeco Large and Mid Cap Fund Reg (G)", "177193497541", 3900.112, 400000, 512340),
    ("Canara Robeco Large and Mid Cap Fund Reg (G)", "177214597312", 2650.884, 260000, 318760),
    ("Mirae Asset Focused Fund Reg (G)", "777323907922", 4980.221, 480000, 561230),
    ("Mirae Asset Large & Midcap Fund Reg (G)", "776894525104", 3120.667, 350000, 429870),
    ("Mirae Asset Large Cap Fund Reg (G)", "775149452517", 5210.334, 520000, 634210),
    ("Nippon India Growth Mid Cap Fund Reg (G)", "477288232358", 1980.129, 310000, 398450),
    ("Parag Parikh Flexi Cap Fund Reg (G)", "116769701", 4780.912, 450000, 578230),
    ("SBI ESG Exclusionary Strategy Fund Reg (G)", "144066101", 2870.556, 320000, 412980),
    ("SBI Focused Fund Reg (G)", "223895051", 3410.221, 350000, 421450),
    ("SBI Large Cap Fund Direct (G)", "144066102", 2980.774, 340000, 398760),
    ("SBI MNC Fund Reg (G)", "144075231", 1780.112, 260000, 312450),
    ("SBI Small Cap Fund Reg (G)", "218114851", 2150.887, 290000, 378120),
    ("Tata Large & Mid Cap Fund Reg (G)", "528100980", 3670.445, 380000, 452340),
]


def _golden_holdings() -> list:
    hs = [_mk_holding("Equity", s, f, u, p, c) for s, f, u, p, c in _GOLDEN_EQUITY]
    hs.append(_mk_holding("Hybrid", "ICICI Pru Equity & Debt Fund", "334455667788", 2210.5, 300000, 356000, cagr=13.0))
    hs.append(_mk_holding("Hybrid", "HSBC Aggressive Hybrid Fund Reg (G)", "345988815", 1980.221, 180000, 214300, cagr=12.4))
    hs.append(_mk_holding("Debt", "ICICI Pru Short Term Fund", "556677889900", 15320.0, 200000, 215000, cagr=6.2))
    hs.append(_mk_holding("Other", "SBI Gold Fund Direct (G)", "144075232", 2851.662, 90000, 118400, cagr=18.9))
    hs.append(_mk_holding("Equity", "Mirae Asset ELSS Tax Saver Fund Reg (G)", "775149452599", 3210.445, 110000, 142300, cagr=15.2))
    return hs


def fixture_golden(mindmap_path=None) -> Fixture:
    holdings = _golden_holdings()
    txns = [
        TransactionSnapshotRow("HDFC Flexicap Fund", "Switch Out", 298000, balance_units=980.234,
                               purchase_amount=250000, suggested_scheme="Parag Parikh Flexicap Fund",
                               purchase_date=date(2024, 1, 15)),
        TransactionSnapshotRow("Parag Parikh Flexicap Fund", "Switch In", None),
        TransactionSnapshotRow("L&T Emerging Businesses Fund", "Redeem", 71500, balance_units=560.912,
                               purchase_amount=80000, purchase_date=date(2026, 3, 1)),
        TransactionSnapshotRow("Kotak Multicap Fund", "Reinvest", 71500),
        TransactionSnapshotRow("Axis Bluechip Fund", "SIP Stop", 10000,
                               suggested_scheme="Mirae Asset Large Cap Fund"),
        TransactionSnapshotRow("Mirae Asset Large Cap Fund", "SIP Start", 10000),
    ]
    return make_context(
        "golden", holdings, txns,
        performance_rows=[
            PerformanceRow("HDFC Flexicap Fund", "out", {"1Y": 14.28, "2Y": 11.2, "3Y": 13.5, "5Y": 12.1,
                                                         "7Y": 11.8, "10Y": 11.55, "Since Launch": 13.0,
                                                         "CY": 9.5, "CY-1": 14.0, "CY-2": 8.0, "CY-3": -5.2, "CY-4": 5.0}),
            PerformanceRow("Parag Parikh Flexicap Fund", "in", {"1Y": 18.9, "2Y": 16.4, "3Y": 19.1, "5Y": 18.0,
                                                                "7Y": 17.2, "10Y": "N/A", "Since Launch": 19.5,
                                                                "CY": 12.1, "CY-1": 21.0, "CY-2": 6.5, "CY-3": 2.3, "CY-4": 15.4}),
        ],
        # Hand-computed, committed alongside the data.
        expected={
            "total_purchase": 9030000,
            "total_current": 10972090,
            "total_gain": 1942090,
            "unique_schemes": 27,
            "equity_exposure_pct": 95.7,   # to 1dp
            "market_cap_bucket_count": 9,
        },
        mindmap_path=mindmap_path,
    )


# ==========================================================================
# B. exemption_blown - three switches, combined gain over the Rs 1.25L cap
# ==========================================================================

def fixture_exemption_blown(mindmap_path=None) -> Fixture:
    # Gains: 80,000 + 90,000 + 60,000 = 2,30,000
    # Taxable: 2,30,000 - 1,25,000 = 1,05,000
    # Tax: 1,05,000 * 12.5% * 1.04 (cess) = 13,650.00
    # Pro-rata: 80/230 -> 4,747.83 | 90/230 -> 5,341.30 | 60/230 -> 3,560.87
    specs = [
        ("Alpha Large Cap Fund", "100000000001", 200000, 280000),   # +80,000
        ("Beta Mid Cap Fund", "100000000002", 300000, 390000),      # +90,000
        ("Gamma Small Cap Fund", "100000000003", 250000, 310000),   # +60,000
    ]
    holdings = [_mk_holding("Equity", s, f, 1000.0, p, c) for s, f, p, c in specs]
    holdings.append(_mk_holding("Debt", "Delta Short Term Fund", "100000000004", 5000.0, 200000, 215000, cagr=6.0))

    txns = []
    for i, (s, _f, p, c) in enumerate(specs, start=1):
        target = f"Target Fund {i}"
        txns.append(TransactionSnapshotRow(s, "Switch Out", c, balance_units=1000.0,
                                           purchase_amount=p, suggested_scheme=target,
                                           purchase_date=date(2024, 1, 15)))
        txns.append(TransactionSnapshotRow(target, "Switch In", None))

    return make_context(
        "exemption_blown", holdings, txns,
        expected={
            "aggregate_tax": 13650.00,
            "allocations": {
                "Alpha Large Cap Fund": 4747.83,
                "Beta Mid Cap Fund": 5341.30,
                "Gamma Small Cap Fund": 3560.87,
            },
            "taxable_gain": 105000.0,
        },
        mindmap_path=mindmap_path,
    )


# ==========================================================================
# C. no_hybrid - zero hybrid holdings
# ==========================================================================

def fixture_no_hybrid(mindmap_path=None) -> Fixture:
    holdings = [
        _mk_holding("Equity", "Solo Large Cap Fund", "200000000001", 1000.0, 400000, 500000),
        _mk_holding("Equity", "Solo Mid Cap Fund", "200000000002", 800.0, 200000, 260000),
        _mk_holding("Debt", "Solo Short Term Fund", "200000000003", 5000.0, 300000, 320000, cagr=5.5),
        _mk_holding("Other", "Solo Gold Fund", "200000000004", 900.0, 100000, 120000, cagr=14.0),
    ]
    # equity 7,60,000 / total 12,00,000 = 63.333...%  (no hybrid term at all)
    return make_context(
        "no_hybrid", holdings, [],
        expected={
            "equity_exposure_pct": 63.33,
            "hybrid_holdings": 0,
            "total_current": 1200000,
        },
        mindmap_path=mindmap_path,
    )


# ==========================================================================
# D. unknown_schemes - four schemes matching no market-cap keyword
# ==========================================================================

def fixture_unknown_schemes(mindmap_path=None) -> Fixture:
    holdings = [
        _mk_holding("Equity", "Quantum Momentum Alpha Fund", "300000000001", 1000.0, 100000, 120000),
        _mk_holding("Equity", "Zephyr Absolute Return Scheme", "300000000002", 1000.0, 150000, 170000),
        _mk_holding("Equity", "Helios Special Situations Plan", "300000000003", 1000.0, 200000, 230000),
        _mk_holding("Equity", "Orion Structured Opportunity", "300000000004", 1000.0, 250000, 280000),
        _mk_holding("Equity", "Axis Large Cap Fund", "300000000005", 1000.0, 300000, 360000),
    ]
    return make_context(
        "unknown_schemes", holdings, [],
        expected={
            "unclassified_count": 4,
            "unclassified_value": 120000 + 170000 + 230000 + 280000,  # 800000
            "large_cap_value": 360000,
        },
        mindmap_path=mindmap_path,
    )


# ==========================================================================
# E. no_dates - ELSS present, purchase dates absent
# ==========================================================================

def fixture_no_dates(mindmap_path=None) -> Fixture:
    holdings = [
        _mk_holding("Equity", "Mirae Asset ELSS Tax Saver Fund Reg (G)", "400000000001", 1000.0, 110000, 142300),
        _mk_holding("Equity", "Axis Large Cap Fund", "400000000002", 1000.0, 300000, 360000),
    ]
    return make_context(
        "no_dates", holdings, [],
        expected={
            "expected_80c_status": (
                "ELSS holdings present but purchase dates not available in uploaded file "
                "- cannot compute FY contribution"
            ),
            "lifetime_cost_must_not_appear": 110000,
        },
        mindmap_path=mindmap_path,
    )


# ==========================================================================
# F. large - 60 holdings across all categories
# ==========================================================================

_LARGE_EQUITY_KINDS = [
    "Large Cap", "Mid Cap", "Small Cap", "Flexi Cap", "Multi Cap",
    "Focused", "Contra", "ESG", "Large and Mid Cap",
]


def fixture_large(mindmap_path=None) -> Fixture:
    holdings = []
    for i in range(48):  # 48 equity
        kind = _LARGE_EQUITY_KINDS[i % len(_LARGE_EQUITY_KINDS)]
        holdings.append(_mk_holding(
            "Equity", f"Testco {kind} Fund Series {i+1} Reg (G)",
            f"{500000000000 + i}", 1000.0 + i, 100000 + i * 1000, 120000 + i * 1300,
        ))
    for i in range(6):   # 6 hybrid
        holdings.append(_mk_holding(
            "Hybrid", f"Testco Aggressive Hybrid Fund Series {i+1}",
            f"{510000000000 + i}", 900.0, 150000, 178000, cagr=11.5,
        ))
    for i in range(4):   # 4 debt
        holdings.append(_mk_holding(
            "Debt", f"Testco Short Duration Fund Series {i+1}",
            f"{520000000000 + i}", 5000.0, 200000, 214000, cagr=5.8,
        ))
    for i in range(2):   # 2 other
        holdings.append(_mk_holding(
            "Other", f"Testco Gold Fund Series {i+1}",
            f"{530000000000 + i}", 800.0, 90000, 110000, cagr=13.2,
        ))
    return make_context(
        "large", holdings, [],
        expected={"holding_count": 60},
        mindmap_path=mindmap_path,
    )


# ==========================================================================
# G. all_loss - every holding under water
# ==========================================================================

def fixture_all_loss(mindmap_path=None) -> Fixture:
    specs = [
        ("Sinking Large Cap Fund", "600000000001", 500000, 420000, -8.5),
        ("Sinking Mid Cap Fund", "600000000002", 300000, 240000, -12.0),
        ("Sinking Flexi Cap Fund", "600000000003", 200000, 176000, -6.2),
    ]
    holdings = [_mk_holding("Equity", s, f, 1000.0, p, c, cagr=g) for s, f, p, c, g in specs]
    holdings.append(_mk_holding("Hybrid", "Sinking Aggressive Hybrid Fund", "600000000004",
                                900.0, 150000, 138000, cagr=-4.1))
    holdings.append(_mk_holding("Debt", "Sinking Short Term Fund", "600000000005",
                                5000.0, 200000, 194000, cagr=-1.8))

    txns = [
        TransactionSnapshotRow("Sinking Mid Cap Fund", "Redeem", 240000, balance_units=1000.0,
                               purchase_amount=300000, purchase_date=date(2026, 3, 1)),
        TransactionSnapshotRow("Sinking Flexi Cap Fund", "Switch Out", 176000, balance_units=1000.0,
                               purchase_amount=200000, suggested_scheme="Recovery Fund",
                               purchase_date=date(2023, 5, 1)),
        TransactionSnapshotRow("Recovery Fund", "Switch In", None),
    ]
    return make_context(
        "all_loss", holdings, txns,
        expected={
            "total_purchase": 1350000,
            "total_current": 1168000,
            "total_gain": -182000,
            "aggregate_tax": 0.0,
            "cagr_is_negative": True,
        },
        mindmap_path=mindmap_path,
    )


# ==========================================================================
# H. dashboard_scale - the real dashboard export's scale
# ==========================================================================
# Stands in for the largest client in the 311-holding / Rs 28.81 Cr file.
# SYNTHETIC: that file is not in this repo, so the holding COUNT and the
# portfolio VALUE are matched to the documented totals and the mix is
# modelled on the other fixtures. Sized to the whole file rather than to
# one client inside it, which upper-bounds any single client and answers
# the question this fixture exists to answer: does the Holdings table
# still paginate, repeat its header, and build in reasonable time at five
# times the previous largest fixture.

_SCALE_EQUITY_KINDS = _LARGE_EQUITY_KINDS
_SCALE_TARGET_VALUE = 288_100_000.0   # Rs 28.81 Cr
_SCALE_HOLDING_COUNT = 311


def fixture_dashboard_scale(mindmap_path=None) -> Fixture:
    equity_n, hybrid_n, debt_n, other_n = 240, 40, 24, 7
    assert equity_n + hybrid_n + debt_n + other_n == _SCALE_HOLDING_COUNT

    holdings = []
    for i in range(equity_n):
        kind = _SCALE_EQUITY_KINDS[i % len(_SCALE_EQUITY_KINDS)]
        holdings.append(_mk_holding(
            "Equity", f"Dashboard {kind} Fund Series {i + 1} Reg (G)",
            f"{600000000000 + i}", 1000.0 + i, 700000 + i * 500, 860000 + i * 700,
        ))
    for i in range(hybrid_n):
        holdings.append(_mk_holding(
            "Hybrid", f"Dashboard Aggressive Hybrid Fund Series {i + 1}",
            f"{610000000000 + i}", 900.0, 500000, 592000, cagr=11.5,
        ))
    for i in range(debt_n):
        holdings.append(_mk_holding(
            "Debt", f"Dashboard Short Duration Fund Series {i + 1}",
            f"{620000000000 + i}", 5000.0, 400000, 428000, cagr=5.8,
        ))
    for i in range(other_n):
        holdings.append(_mk_holding(
            "Other", f"Dashboard Gold Fund Series {i + 1}",
            f"{630000000000 + i}", 800.0, 300000, 366000, cagr=13.2,
        ))

    # Scale every holding by one factor so the portfolio lands on the
    # documented Rs 28.81 Cr without disturbing the per-holding mix. The
    # gain and return fields are recomputed from the scaled figures rather
    # than scaled independently, so the invariants still hold exactly.
    raw_total = sum(h.current_value for h in holdings)
    factor = _SCALE_TARGET_VALUE / raw_total
    for h in holdings:
        h.purchase_value = round(h.purchase_value * factor, 2)
        h.current_value = round(h.current_value * factor, 2)
        h.gain = round(h.current_value - h.purchase_value, 2)
        h.absolute_return_pct = round(h.gain / h.purchase_value * 100, 2) if h.purchase_value else None

    return make_context(
        "dashboard_scale", holdings, [],
        expected={"holding_count": _SCALE_HOLDING_COUNT},
        mindmap_path=mindmap_path,
    )


ALL_FIXTURE_BUILDERS = {
    "A_golden": fixture_golden,
    "B_exemption_blown": fixture_exemption_blown,
    "C_no_hybrid": fixture_no_hybrid,
    "D_unknown_schemes": fixture_unknown_schemes,
    "E_no_dates": fixture_no_dates,
    "F_large": fixture_large,
    "G_all_loss": fixture_all_loss,
    "H_dashboard_scale": fixture_dashboard_scale,
}
