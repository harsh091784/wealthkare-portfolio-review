"""
pipeline/risk_profile.py

Risk profile computation for the WC Securities / Wealthkare Portfolio
Review pipeline.

CRITICAL design rule (this was a production bug previously): the risk
profile shown in a report must be computed FRESH, every single run, from
the client's *actual current* equity:debt split. It is never taken from a
client-declared questionnaire answer, and the profile label/description is
never hardcoded or cached anywhere - a stale "Aggressive" label must not
survive a portfolio whose current allocation now classifies as Moderate.

Because the profile is INFERRED from current allocation, there is no
questionnaire-derived target allocation anywhere in this pipeline. A
portfolio therefore cannot be "off target", "misaligned", or "drifted
from its band" - the allocation IS what defines the band. Report copy
must never imply otherwise.

Look-through logic
-------------------
Hybrid/blended schemes don't sit 100% in either bucket. We attribute a
fixed equity weight to each recognised scheme type before summing:

    Aggressive Hybrid              75% equity
    Balanced Advantage / BAF       50% equity
    Multi-Asset Allocation         35% equity
    Arbitrage Fund                  0% equity  (goes to Debt bucket)
    Debt / Liquid / Money Market    0% equity  (goes to Debt bucket)
    Gold / SGB / Gold ETF           0% equity  (goes to Others bucket)

Plain equity schemes (Large Cap, Mid Cap, Small Cap, Flexi Cap, Focused,
Sectoral/Thematic, Index, ELSS, Value/Contra, ...) are treated as 100%
equity and also drive the equity sub-allocation breakdown table.

Anything that matches none of the above known patterns is NOT guessed -
its value is excluded from the equity/debt/others totals, logged as a
warning, and reported separately as "unclassified" so it's visible rather
than silently mis-weighted.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional

# --------------------------------------------------------------------------
# Data model
# --------------------------------------------------------------------------

@dataclass
class RiskHolding:
    """Minimal input needed per holding to run the risk-profile calc.

    `fund_type` is a free-text scheme/category description (e.g. "Large
    Cap Fund", "Aggressive Hybrid Fund", "Balanced Advantage Fund",
    "Liquid Fund", "Gold ETF"). It does not need to be an exact scheme
    name - classification is keyword-based on this field.
    """
    scheme: str
    fund_type: str
    current_value: float
    member: Optional[str] = None


@dataclass
class AllocationRow:
    label: str
    value: float
    pct_of_bucket: float  # % of the bucket this row belongs to (e.g. % of total equity)


@dataclass
class RiskProfileResult:
    equity_value: float = 0.0
    debt_value: float = 0.0
    others_value: float = 0.0
    unclassified_value: float = 0.0
    total_classified_value: float = 0.0
    equity_pct: Optional[float] = None
    debt_pct: Optional[float] = None
    others_pct: Optional[float] = None
    profile: Optional[str] = None
    # Equity-exposure RANGE describing the band this profile covers, e.g.
    # "75-100% equity exposure". Deliberately NOT an "80/20"-style ratio:
    # WC infers the profile from current allocation, so there is no target
    # split, and a ratio printed beside the client's own figure reads as
    # if it were their split.
    band_definition: Optional[str] = None
    description: Optional[str] = None
    equity_sub_allocation: list[AllocationRow] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


# --------------------------------------------------------------------------
# Classification rules
# --------------------------------------------------------------------------

# (keywords, equity_weight, non_equity_bucket, display_label)
# Checked in order - more specific patterns MUST come before broader ones
# (e.g. "balanced advantage" before a bare "balanced").
HYBRID_DEBT_GOLD_RULES: list[tuple[list[str], float, str, str]] = [
    (["gold etf", "gold fund", "sgb", "sovereign gold"], 0.0, "Others", "Gold / SGB / Gold ETF"),
    (["aggressive hybrid"], 0.75, "Debt", "Aggressive Hybrid"),
    (["balanced advantage", "baf"], 0.50, "Debt", "Balanced Advantage / BAF"),
    (["multi-asset", "multi asset"], 0.35, "Debt", "Multi-Asset Allocation"),
    (["arbitrage"], 0.0, "Debt", "Arbitrage Fund"),
    (
        ["liquid", "money market", "overnight", "ultra short", "low duration",
         "short duration", "corporate bond", "banking & psu", "banking and psu",
         "gilt", "credit risk", "dynamic bond", "debt fund", "income fund",
         "medium duration", "long duration"],
        0.0, "Debt", "Debt / Liquid / Money Market",
    ),
]

# (keyword, display sub-category label) - pure equity, 100% weight.
EQUITY_SUBCATEGORY_RULES: list[tuple[str, str]] = [
    ("large & mid cap", "Large & Mid Cap"),
    ("large and mid cap", "Large & Mid Cap"),
    ("large cap", "Large Cap"),
    ("mid cap", "Mid Cap"),
    ("small cap", "Small Cap"),
    ("flexi cap", "Flexi Cap"),
    ("multi cap", "Multi Cap"),
    ("focused", "Focused Fund"),
    ("sectoral", "Sectoral / Thematic"),
    ("thematic", "Sectoral / Thematic"),
    ("index", "Index Fund"),
    ("elss", "ELSS / Tax Saver"),
    ("value", "Value / Contra"),
    ("contra", "Value / Contra"),
    ("dividend yield", "Dividend Yield"),
]

# Equity % bands, nearest-match. (profile name, equity %, debt %)
#
# NOTE on the second/third numbers: these are the band ANCHOR points used
# for nearest-match classification, NOT a target split for any client. WC
# infers a client's risk profile purely from their current allocation -
# there is no questionnaire-derived target to compare against, so a
# portfolio can never be "off target" or "drifted". Anything user-facing
# must therefore describe the band as a RANGE of equity exposure
# (see band_definition_for_profile) and never print an "80/20"-style
# ratio next to the client's own figure, where it reads as their split.
RISK_BANDS: list[tuple[str, int, int]] = [
    ("Aggressive", 80, 20),
    ("Moderately Aggressive", 70, 30),
    ("Moderate", 60, 40),
    ("Moderately Conservative", 50, 50),
    ("Conservative", 40, 60),
]


def band_definition_for_profile(profile: str) -> Optional[str]:
    """Human-readable equity-exposure RANGE for a profile, e.g.
    'Aggressive' -> '75-100% equity exposure'.

    Derived from the actual nearest-match boundaries between adjacent
    RISK_BANDS anchors (the midpoints), so the printed range is exactly
    the range that classification really uses - not a hand-written
    approximation that could quietly disagree with the code.
    """
    anchors = sorted(RISK_BANDS, key=lambda b: b[1], reverse=True)  # highest equity first
    for i, (name, equity_anchor, _debt_anchor) in enumerate(anchors):
        if name != profile:
            continue
        upper = 100.0 if i == 0 else (equity_anchor + anchors[i - 1][1]) / 2
        lower = 0.0 if i == len(anchors) - 1 else (equity_anchor + anchors[i + 1][1]) / 2
        return f"{lower:g}-{upper:g}% equity exposure"
    return None

PROFILE_DESCRIPTIONS: dict[str, str] = {
    "Aggressive": (
        "You have emerged as an investor with the Aggressive risk profile. "
        "The primary goal of yours is a significant level of growth in "
        "income. It allows you to pursue higher risk, with the potential "
        "to earn higher return."
    ),
    "Moderately Aggressive": (
        "You have emerged as a Moderately Aggressive investor. You seek "
        "above-average growth while maintaining some buffer against "
        "extreme market swings."
    ),
    "Moderate": (
        "You have emerged as a Moderate investor. You balance growth and "
        "stability, accepting average market risk for average long-term "
        "returns."
    ),
    "Moderately Conservative": (
        "You have emerged as a Moderately Conservative investor. Capital "
        "preservation with modest growth is your primary objective."
    ),
    "Conservative": (
        "You have emerged as a Conservative investor. Preserving capital "
        "takes priority over growth, with a preference for stable, "
        "lower-risk instruments."
    ),
}


# --------------------------------------------------------------------------
# Classification
# --------------------------------------------------------------------------

def _classify_fund_type(fund_type: str) -> Optional[tuple[float, str, str]]:
    """Returns (equity_weight, non_equity_bucket, display_label) or None if
    the fund_type doesn't match any known pattern."""
    text = (fund_type or "").strip().lower()
    if not text:
        return None

    for keywords, weight, bucket, label in HYBRID_DEBT_GOLD_RULES:
        if any(kw in text for kw in keywords):
            return weight, bucket, label

    for kw, label in EQUITY_SUBCATEGORY_RULES:
        if kw in text:
            return 1.0, "Debt", label  # bucket unused when weight == 1.0

    return None


# --------------------------------------------------------------------------
# Core computation
# --------------------------------------------------------------------------

def compute_risk_profile(holdings: list[RiskHolding]) -> RiskProfileResult:
    result = RiskProfileResult()

    equity_value = 0.0
    debt_value = 0.0
    others_value = 0.0
    unclassified_value = 0.0

    # scheme "display label" -> summed value, for the equity sub-allocation
    # breakdown table. Both pure-equity holdings and the equity-weighted
    # look-through portion of hybrid schemes are included, each under its
    # own label, so the table is transparent about where the equity is
    # actually coming from.
    equity_breakdown: dict[str, float] = {}

    for h in holdings:
        if h.current_value is None:
            result.warnings.append(
                f"'{h.scheme}': current_value is missing - excluded from risk profile calc."
            )
            continue

        classified = _classify_fund_type(h.fund_type)
        if classified is None:
            unclassified_value += h.current_value
            result.warnings.append(
                f"'{h.scheme}' (fund_type='{h.fund_type}'): no matching classification rule - "
                f"excluded from equity/debt totals rather than guessed. Value: {h.current_value:,.2f}"
            )
            continue

        weight, non_equity_bucket, label = classified
        equity_portion = h.current_value * weight
        non_equity_portion = h.current_value - equity_portion

        equity_value += equity_portion
        if equity_portion > 0:
            equity_breakdown[label] = equity_breakdown.get(label, 0.0) + equity_portion

        if non_equity_portion > 0:
            if non_equity_bucket == "Others":
                others_value += non_equity_portion
            else:
                debt_value += non_equity_portion

    total_classified_value = equity_value + debt_value + others_value

    result.equity_value = round(equity_value, 2)
    result.debt_value = round(debt_value, 2)
    result.others_value = round(others_value, 2)
    result.unclassified_value = round(unclassified_value, 2)
    result.total_classified_value = round(total_classified_value, 2)

    if total_classified_value <= 0:
        result.warnings.append(
            "No classified holdings with positive value - risk profile cannot be computed."
        )
        return result

    equity_pct = (equity_value / total_classified_value) * 100
    debt_pct = (debt_value / total_classified_value) * 100
    others_pct = (others_value / total_classified_value) * 100

    result.equity_pct = round(equity_pct, 2)
    result.debt_pct = round(debt_pct, 2)
    result.others_pct = round(others_pct, 2)

    # Nearest-band match on equity % - recomputed fresh every call, never
    # cached or hardcoded.
    profile_name, band_equity, band_debt = min(
        RISK_BANDS, key=lambda band: abs(equity_pct - band[1])
    )
    result.profile = profile_name
    result.band_definition = band_definition_for_profile(profile_name)
    result.description = PROFILE_DESCRIPTIONS[profile_name]

    # Equity sub-allocation breakdown table, sorted largest first.
    result.equity_sub_allocation = [
        AllocationRow(
            label=label,
            value=round(value, 2),
            pct_of_bucket=round((value / equity_value) * 100, 2) if equity_value > 0 else 0.0,
        )
        for label, value in sorted(equity_breakdown.items(), key=lambda kv: kv[1], reverse=True)
    ]

    if unclassified_value > 0:
        result.warnings.append(
            f"Total unclassified value excluded from risk profile: {unclassified_value:,.2f} "
            f"(see per-holding warnings above)."
        )

    return result


# --------------------------------------------------------------------------
# Headline equity exposure - ONE definition, used everywhere a report needs
# to state "your equity %" as a single figure.
# --------------------------------------------------------------------------
#
# This was a real bug: the asset-allocation pie showed "Equity 91.0%" (the
# literal category=="Equity" holdings only) while the risk-profile gauge
# showed "You: 95.6%" (compute_risk_profile()'s full per-subtype look-
# through, which folds in 75% of Aggressive Hybrid holdings). Same client,
# same report, two different numbers with two different denominators and
# no way for a reader to know that. compute_headline_equity_exposure_pct()
# is now the ONLY function either chart is allowed to call for this
# number, so they can't drift apart again.
#
# The definition here (direct equity + 75% of Hybrid, over the grand
# total) is deliberately simpler than compute_risk_profile()'s per-subtype
# weighting (Aggressive Hybrid 75% / BAF 50% / Multi-Asset 35% / ...) -
# that granular system still drives which risk BAND the client lands in
# (compute_risk_profile().profile), which this function does not change.
# This is purely the headline display number.

def compute_headline_equity_exposure_pct(holdings) -> Optional[float]:
    """holdings: any iterable of objects with `.category` (str, expected to
    contain "Equity" / "Hybrid" / etc.) and `.current_value` (float) -
    i.e. pipeline.parser.Holding objects. Returns None if the grand total
    is zero (can't compute a percentage of nothing)."""
    equity_value = sum(
        h.current_value for h in holdings
        if h.category == "Equity" and h.current_value is not None
    )
    hybrid_value = sum(
        h.current_value for h in holdings
        if h.category == "Hybrid" and h.current_value is not None
    )
    grand_total = sum(h.current_value for h in holdings if h.current_value is not None)
    if not grand_total:
        return None
    return (equity_value + 0.75 * hybrid_value) / grand_total * 100


HEADLINE_EQUITY_EXPOSURE_FOOTNOTE = (
    "Equity exposure = direct equity holdings + 75% look-through on Hybrid "
    "holdings, as a % of the total portfolio value. This is the same figure "
    "used for the Risk Profile gauge."
)


# --------------------------------------------------------------------------
# Market-cap classification - equity holdings only, scheme-name based.
# --------------------------------------------------------------------------
#
# Deliberately separate from EQUITY_SUBCATEGORY_RULES/HYBRID_DEBT_GOLD_RULES
# above: this table classifies PURE EQUITY holdings into market-cap-size
# buckets for the "Equity Sub-Allocation by Market Cap" table under Current
# Asset Allocation. Hybrid funds never appear here - "how hybrid a fund is"
# and "what market cap it invests in" are different questions, and hybrid
# look-through belongs in the asset-allocation pie, not this table.
#
# This is an explicit, ordered keyword rule table (checked most-specific
# first, e.g. "large & mid cap" before a bare "large cap") rather than a
# literal per-scheme dictionary - AMC scheme names change/launch constantly,
# so a hardcoded scheme->category dict would be stale within a quarter.
# Indian AMFI scheme names near-universally spell out their SEBI category
# in the name itself ("...Large Cap Fund", "...Flexi Cap Fund", "...ELSS
# Tax Saver..."), which is what makes keyword matching reliable here. Any
# scheme name that doesn't match a known keyword is "Unclassified" and
# reported as a warning - NEVER silently defaulted to Large Cap.

MARKET_CAP_RULES: list[tuple[list[str], str]] = [
    (["large & mid cap", "large and mid cap"], "Large & Mid Cap"),
    (["large cap", "bluechip", "blue chip"], "Large Cap"),
    (["mid cap"], "Mid Cap"),
    (["small cap", "emerging businesses"], "Small Cap"),
    (["flexi cap", "flexicap"], "Flexi Cap"),
    (["multi cap", "multicap"], "Multi Cap"),
    (["focused"], "Focused"),
    (["contra", "value"], "Value/Contra"),
    (
        ["sectoral", "thematic", "esg", "mnc", "consumption", "infrastructure",
         "banking", "pharma", "technology", "digital"],
        "Sectoral/Thematic",
    ),
    (["elss", "tax saver"], "ELSS"),
]


def classify_scheme_market_cap(scheme_name: str) -> str:
    """Returns one of the MARKET_CAP_RULES labels, or 'Unclassified' if the
    scheme name doesn't contain a recognised market-cap keyword."""
    text = (scheme_name or "").lower()
    # Normalize common no-space variants ("Midcap" / "Smallcap" / "Largecap")
    # so they still match the space-separated keywords above. "flexicap"
    # and "multicap" are matched directly instead, since those ARE the
    # standard no-space spellings in real scheme names.
    text = text.replace("midcap", "mid cap").replace("smallcap", "small cap").replace("largecap", "large cap")
    for keywords, label in MARKET_CAP_RULES:
        if any(kw in text for kw in keywords):
            return label
    return "Unclassified"


def compute_equity_market_cap_breakdown(
    holdings, overrides: Optional[dict] = None
) -> tuple[list[AllocationRow], list[str]]:
    """Market-cap sub-allocation for PURE EQUITY holdings only (category ==
    "Equity"). Returns (rows, warnings) - rows sorted largest first,
    warnings listing every scheme that came back "Unclassified" (fix #3:
    never silently defaulted to Large Cap).

    holdings: any iterable of objects with `.category`, `.scheme`,
    `.current_value` - i.e. pipeline.parser.Holding objects.

    overrides: optional {scheme: label} decided by a HUMAN for schemes the
    keyword rules could not place. Purely additive - omitting it leaves
    the classification exactly as it was. A scheme is only overridable if
    the rules failed on it; an override is never allowed to silently
    contradict a rule that DID match, because that would let a stale
    manual choice outvote the scheme's own name.
    """
    overrides = overrides or {}
    equity_holdings = [h for h in holdings if h.category == "Equity" and h.current_value]
    total_equity_value = sum(h.current_value for h in equity_holdings)

    warnings: list[str] = []
    buckets: dict[str, float] = {}
    for h in equity_holdings:
        label = classify_scheme_market_cap(h.scheme)
        if label == "Unclassified" and h.scheme in overrides:
            label = overrides[h.scheme]
        elif label == "Unclassified":
            warnings.append(
                f"'{h.scheme}': scheme name did not match any known market-cap keyword - "
                f"bucketed as Unclassified rather than defaulted to Large Cap."
            )
        buckets[label] = buckets.get(label, 0.0) + h.current_value

    rows = [
        AllocationRow(
            label=label,
            value=round(value, 2),
            pct_of_bucket=round(value / total_equity_value * 100, 2) if total_equity_value else 0.0,
        )
        for label, value in sorted(buckets.items(), key=lambda kv: kv[1], reverse=True)
    ]
    return rows, warnings


# --------------------------------------------------------------------------
# Self-test (dummy holdings, no external files required)
# --------------------------------------------------------------------------

def _run_self_test() -> None:
    print("=== pipeline/risk_profile.py self-test ===\n")

    # --- Case 1: a mixed portfolio expected to land near Moderately Aggressive ---
    holdings_1 = [
        RiskHolding("Axis Bluechip Fund", "Large Cap Fund", 200000),
        RiskHolding("Kotak Emerging Equity Fund", "Mid Cap Fund", 100000),
        RiskHolding("Nippon India Small Cap Fund", "Small Cap Fund", 50000),
        RiskHolding("ICICI Pru Equity & Debt Fund", "Aggressive Hybrid Fund", 80000),
        RiskHolding("HDFC Balanced Advantage Fund", "Balanced Advantage Fund", 60000),
        RiskHolding("Kotak Multi Asset Allocator", "Multi-Asset Allocation Fund", 30000),
        RiskHolding("Edelweiss Arbitrage Fund", "Arbitrage Fund", 40000),
        RiskHolding("HDFC Liquid Fund", "Liquid Fund", 50000),
        RiskHolding("SBI Gold ETF", "Gold ETF", 20000),
        RiskHolding("Weird Structured Note Series 3", "Structured Note Series 3", 15000),  # unclassifiable on purpose
    ]

    result_1 = compute_risk_profile(holdings_1)

    print("--- Case 1: mixed portfolio ---")
    print(f"Equity value:   {result_1.equity_value:,.2f}  ({result_1.equity_pct}%)")
    print(f"Debt value:     {result_1.debt_value:,.2f}  ({result_1.debt_pct}%)")
    print(f"Others value:   {result_1.others_value:,.2f}  ({result_1.others_pct}%)")
    print(f"Unclassified:   {result_1.unclassified_value:,.2f}")
    print(f"Total classified: {result_1.total_classified_value:,.2f}")
    print(f"Profile: {result_1.profile}  (band: {result_1.band_definition})")
    print(f"Description: {result_1.description}")
    print("\nEquity sub-allocation:")
    for row in result_1.equity_sub_allocation:
        print(f"  {row.label:<32} {row.value:>12,.2f}  ({row.pct_of_bucket}% of equity)")
    print(f"\nWarnings ({len(result_1.warnings)}):")
    for w in result_1.warnings:
        print(f"  - {w}")

    assert result_1.profile is not None
    assert result_1.unclassified_value == 15000
    assert any("Structured Note" in w for w in result_1.warnings)

    # --- Case 2: heavily debt-weighted portfolio -> should land Conservative ---
    holdings_2 = [
        RiskHolding("HDFC Short Term Debt Fund", "Debt Fund", 90000),
        RiskHolding("ICICI Pru Liquid Fund", "Liquid Fund", 60000),
        RiskHolding("Axis Bluechip Fund", "Large Cap Fund", 30000),
    ]
    result_2 = compute_risk_profile(holdings_2)
    print("\n--- Case 2: debt-heavy portfolio ---")
    print(f"Equity %: {result_2.equity_pct}%  Profile: {result_2.profile}  (band: {result_2.band_definition})")
    assert result_2.profile == "Conservative", f"Expected Conservative, got {result_2.profile}"

    # --- Case 3: fully equity portfolio -> should land Aggressive ---
    holdings_3 = [
        RiskHolding("Parag Parikh Flexi Cap Fund", "Flexi Cap Fund", 100000),
        RiskHolding("Motilal Oswal Nasdaq 100 FOF", "Sectoral/Thematic Fund", 20000),
    ]
    result_3 = compute_risk_profile(holdings_3)
    print("\n--- Case 3: fully equity portfolio ---")
    print(f"Equity %: {result_3.equity_pct}%  Profile: {result_3.profile}  (band: {result_3.band_definition})")
    assert result_3.profile == "Aggressive", f"Expected Aggressive, got {result_3.profile}"

    # --- Case 4: empty / all-unclassified -> no profile, just a warning ---
    holdings_4 = [RiskHolding("Mystery Fund", "Unknown Category XYZ", 10000)]
    result_4 = compute_risk_profile(holdings_4)
    print("\n--- Case 4: fully unclassified portfolio ---")
    print(f"Profile: {result_4.profile}  Warnings: {result_4.warnings}")
    assert result_4.profile is None
    assert result_4.unclassified_value == 10000

    # --- compute_headline_equity_exposure_pct / classify_scheme_market_cap
    # / compute_equity_market_cap_breakdown: operate on raw holdings
    # (category + scheme + current_value), not RiskHolding/fund_type. ---
    from dataclasses import dataclass as _dc

    @_dc
    class _FakeHolding:
        category: str
        scheme: str
        current_value: float

    fake_holdings = [
        _FakeHolding("Equity", "Axis Large Cap Fund", 100000),
        _FakeHolding("Equity", "HDFC Mid Cap Fund", 50000),
        _FakeHolding("Equity", "Weird Momentum Fund", 20000),   # no market-cap keyword -> Unclassified
        _FakeHolding("Hybrid", "ICICI Pru Equity & Debt Fund", 80000),
        _FakeHolding("Debt", "ICICI Pru Short Term Fund", 50000),
    ]

    print("\n--- compute_headline_equity_exposure_pct ---")
    exposure = compute_headline_equity_exposure_pct(fake_holdings)
    # (100000+50000+20000 direct equity) + 0.75*80000 hybrid = 170000+60000=230000
    # over grand total (100000+50000+20000+80000+50000=300000) = 76.67%
    expected_exposure = (170000 + 0.75 * 80000) / 300000 * 100
    print(f"Equity exposure: {exposure:.2f}%  (expected {expected_exposure:.2f}%)")
    assert abs(exposure - expected_exposure) < 0.01

    print("\n--- classify_scheme_market_cap ---")
    test_cases = [
        ("Axis Large Cap Fund", "Large Cap"),
        ("HDFC Mid Cap Fund", "Mid Cap"),
        ("Canara Robeco Large and Mid Cap Fund Reg (G)", "Large & Mid Cap"),
        ("Kotak Large & Midcap Fund Reg (G)", "Large & Mid Cap"),  # no-space "Midcap" variant
        ("SBI Small Cap Fund", "Small Cap"),
        ("HDFC Flexicap Fund", "Flexi Cap"),
        ("Nippon India Multicap Fund", "Multi Cap"),
        ("Axis Focused Fund", "Focused"),
        ("Tata Digital India Fund", "Sectoral/Thematic"),  # "digital" AMFI alias
        ("Mirae Asset ELSS Tax Saver Fund", "ELSS"),
        ("Axis Bluechip Fund", "Large Cap"),  # "bluechip" synonym
        ("Invesco India Contra Fund", "Value/Contra"),
        ("ICICI Pru Value Discovery Fund", "Value/Contra"),
        ("SBI ESG Exclusionary Strategy Fund", "Sectoral/Thematic"),
        ("SBI MNC Fund", "Sectoral/Thematic"),
        ("SBI Consumption Opportunities Fund", "Sectoral/Thematic"),
        ("ICICI Pru Infrastructure Fund", "Sectoral/Thematic"),
        ("HDFC Banking & Financial Services Fund", "Sectoral/Thematic"),
        ("Nippon India Pharma Fund", "Sectoral/Thematic"),
        ("ICICI Pru Technology Fund", "Sectoral/Thematic"),
        ("L&T Emerging Businesses Fund", "Small Cap"),  # AMFI alias, not a literal "small cap" match
        ("Weird Momentum Fund", "Unclassified"),  # never silently defaulted to Large Cap
    ]
    for scheme, expected in test_cases:
        actual = classify_scheme_market_cap(scheme)
        status = "OK" if actual == expected else "MISMATCH"
        print(f"  [{status}] {scheme!r:52} -> {actual} (expected {expected})")
        assert actual == expected, f"{scheme!r}: expected {expected}, got {actual}"

    print("\n--- compute_equity_market_cap_breakdown ---")
    breakdown_rows, breakdown_warnings = compute_equity_market_cap_breakdown(fake_holdings)
    for row in breakdown_rows:
        print(f"  {row.label:<16} {row.value:>10,.2f}  ({row.pct_of_bucket}% of equity)")
    print(f"  Warnings: {breakdown_warnings}")
    # Hybrid must NOT appear anywhere in this table (fix #3's explicit "remove
    # Aggressive Hybrid from the market-cap table" requirement).
    assert all(row.label != "Aggressive Hybrid" for row in breakdown_rows)
    assert any(row.label == "Unclassified" and row.value == 20000 for row in breakdown_rows)
    assert any("Weird Momentum Fund" in w for w in breakdown_warnings)

    print("\nAll self-test assertions passed.")


if __name__ == "__main__":
    _run_self_test()
