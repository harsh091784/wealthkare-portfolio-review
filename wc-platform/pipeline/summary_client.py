"""
pipeline/summary_client.py

The ONLY LLM call in this pipeline. Everything else - parsing, risk
profiling, tax, allocation, pagination - is deterministic and stays that
way. This module narrates figures that have already been computed; it
never computes one.

Three layers guard that boundary:

  1. INPUT   - build_summary_input() hands the model a dict of computed
               scalars and short structured records. Never raw holdings,
               never free text (no Director's letter, no RM notes, no
               prior report prose). If a figure isn't in this dict, the
               model has no way to know it.

  2. OUTPUT  - validate_summary() re-reads the generated text and fails it
               if any figure in the prose is absent from the input dict,
               if it uses banned target-allocation / forward-return
               language, or if it names a fund that was never supplied.
               One retry with the failure reason appended, then fallback.

  3. GATE    - the generated text is a DRAFT. ClientSummary.approved
               defaults to False and docx_builder.build_report() refuses
               to build while it is False, so nothing a model wrote
               reaches a client without a human saying so.

The deterministic fallback (render_fallback_summary) is used when the API
is unreachable, times out, or fails validation twice. It reads acceptably
rather than brilliantly - the point is that the report ALWAYS builds,
including with no network at all.
"""

from __future__ import annotations

import math
import re
from dataclasses import dataclass, field
from datetime import date
from typing import Any, Optional

from pipeline.chart_gen import format_inr

# The report's OWN formatters, imported rather than reimplemented. A second
# copy of "how a percentage is printed" is how the summary came to say
# 95.66% while the risk gauge and the allocation footnote both said 95.7%
# - two figures for one quantity in the same client document.
from pipeline.docx_builder import EQUITY_EXPOSURE_DECIMALS, _format_pct

# --------------------------------------------------------------------------
# Model configuration
# --------------------------------------------------------------------------

MODEL = "claude-sonnet-4-6"
MAX_TOKENS = 1000
TEMPERATURE = 0.3          # factual narration, not creative writing

# Hard wall-clock bound on a single attempt. The SDK retries timeouts by
# default, which would make the real ceiling timeout x (max_retries + 1);
# API_MAX_RETRIES = 0 keeps 15s meaning 15s. Retrying a *validation*
# failure is handled explicitly below and is a different thing.
API_TIMEOUT_SECONDS = 15.0
API_MAX_RETRIES = 0

MAX_ATTEMPTS = 2           # first attempt + one retry carrying the reason


# --------------------------------------------------------------------------
# System prompt
# --------------------------------------------------------------------------
# Each constraint here has a matching check in validate_summary() wherever
# a check is mechanically possible. The prompt is what makes good output
# likely; the validator is what makes bad output unrenderable.

SYSTEM_PROMPT = """You are writing the Client Summary section of a portfolio review report for WC Securities Pvt Ltd (Wealthkare), an AMFI-registered mutual fund distributor in India.

You will be given a JSON object of figures that have already been computed from the client's actual portfolio. Narrate those figures. You are a narrator, not an analyst.

Rules, all of which are checked automatically before your text is accepted:

1. Every number in your output must appear in the input JSON. Do not compute, derive, round differently, or infer any figure. If you want to state a figure that is not in the input, leave it out instead. Where a field ends in "_display", write that string EXACTLY as given - it is the same string the rest of the report prints, and re-rounding it would put two different figures for one quantity in front of the same client.

2. Do not state or imply that the portfolio is aligned, misaligned, or drifting relative to a target. The risk profile is INFERRED from the current allocation - there is no target to compare against. State the profile and the exposure as facts.

3. Do not predict returns, promise outcomes, or use "will" about future performance. Past performance language only.

4. Do not recommend anything that is not already in the transaction list or the Things To Do list. No new fund suggestions, no new actions.

5. Do not mention data the report says is unavailable (emergency fund, insurance cover) beyond noting that it should be discussed. Tax deductions under 80C, 80D and 80CCD are NOT part of this report - do not mention them at all.

6. Do not add temporal or scope qualifiers that the input does not support. Specifically: do not write "since inception", "over the past year", or any timeframe unless a corresponding date field is present. Do not write "all schemes", "every holding", "none of the funds", or any universal claim about the portfolio - you are shown only the top and bottom performers, never the full list. Describe only what the input explicitly contains.

7. Do not explain WHY a figure is what it is unless the input states the reason. Report the value and stop. If tax is zero, the field tax_zero_reason gives the supported reason - use that wording or say nothing about the cause. Do not write "as the gains fall within the applicable thresholds" or any similar explanation you have inferred rather than been given.

8. This report is written BY this firm TO the client. Write in the first person plural - "us", "we", "your next review". Never refer to "your advisor" or "your financial advisor" in the third person; that is us.

9. State the earliest Things To Do deadline explicitly, with what is due by then. The field earliest_deadline gives it. An action item the client does not see is an action item nobody does.

10. Write 3 to 4 short paragraphs in plain language, second person ("your portfolio"), Indian English. No jargon without explanation. Write only the summary text - no heading, no preamble, no sign-off.
"""


# --------------------------------------------------------------------------
# 1. INPUT - computed values only
# --------------------------------------------------------------------------

def _round_or_none(value: Optional[float], places: int = 2) -> Optional[float]:
    return None if value is None else round(float(value), places)


def _return_ranked_holdings(holdings: list, count: int = 3) -> tuple[list, list]:
    """Top-N holdings by absolute return %, and the worst N holdings that
    are ACTUALLY DOWN.

    Holdings with no return figure are excluded from both ends rather
    than sorted as if they returned zero - a missing return is unknown,
    not flat, and ranking it as flat would put it in the middle of a list
    the summary then describes as "your best and worst".

    The bottom list is capped to negative returns: fewer than N negative
    holdings returns however many there are, and none returns an empty
    list. Taking the numerically lowest three regardless of sign made the
    summary introduce a fund that had GAINED 14.95% as being "on the
    other end" - accurate as a ranking, wrong as a sentence a client
    reads. An empty list means there is nothing to say about
    underperformers, and the summary should say nothing.
    """
    ranked = [
        h for h in holdings
        if getattr(h, "absolute_return_pct", None) is not None and getattr(h, "scheme", None)
    ]
    ranked.sort(key=lambda h: h.absolute_return_pct, reverse=True)

    def record(h) -> dict:
        return {
            "scheme": h.scheme,
            "return_pct_display": _format_pct(h.absolute_return_pct),
            "current_value_rupees": _round_or_none(h.current_value, 0),
            "current_value_display": format_inr(h.current_value)
            if h.current_value is not None else None,
        }

    top = [record(h) for h in ranked[:count]]
    losing = [h for h in ranked if h.absolute_return_pct < 0]
    bottom = [record(h) for h in sorted(losing, key=lambda h: h.absolute_return_pct)[:count]]
    return top, bottom


def _equity_exposure_display(holdings: list) -> Optional[str]:
    """The headline equity figure, printed at the SAME precision as the
    risk gauge and the allocation footnote. Both call the same computation
    (compute_headline_equity_exposure_pct) and print it through
    EQUITY_EXPOSURE_DECIMALS; this reads both from their source rather
    than re-deriving either."""
    from pipeline.risk_profile import compute_headline_equity_exposure_pct
    value = compute_headline_equity_exposure_pct(holdings)
    if value is None:
        return None
    return f"{value:.{EQUITY_EXPOSURE_DECIMALS}f}%"


def _tax_zero_reason(tax) -> Optional[str]:
    """Why the computed tax is zero, decided deterministically from what
    the tax module already computed - never inferred by the model.

    Without this the model supplied its own reason and got it wrong: on an
    all-loss portfolio it wrote that tax was zero "as the gains fall
    within the applicable thresholds", when there were no gains at all.
    Every figure in that sentence was correct and the causal claim was
    invented. None means the tax is non-zero and needs no explanation.
    """
    summary = tax.summary
    if round(summary.total_computed_tax, 2) != 0:
        return None
    if not tax.holdings:
        return "no taxable transactions proposed"
    gross = (
        summary.equity_ltcg_gross_gain
        + summary.equity_stcg_gross_gain
        + summary.non_equity_ltcg_gross_gain
        + summary.non_equity_stcg_gross_gain
    )
    if gross <= 0:
        return "no taxable gains - losses only"
    if summary.equity_ltcg_exemption_applied > 0:
        return "gains within annual exemption"
    return "no tax computed on the proposed transactions"


def _earliest_deadline(things_to_do: list) -> Optional[dict]:
    """The soonest action item, surfaced as its own field so the summary
    states it rather than picking whichever item reads best. It is the one
    with the least time left, and an action item the client never sees is
    an action item nobody does."""
    from pipeline.docx_builder import _parse_deadline
    if not things_to_do:
        return None
    item = min(things_to_do, key=lambda i: _parse_deadline(i.deadline))
    return {
        "deadline": item.deadline,
        "action": item.action,
        "scheme": item.scheme,
        "what_to_do": item.what_to_do,
    }


def build_summary_input(ctx) -> dict:
    """The complete, and only, payload the model sees.

    Deliberately excluded: ctx.holdings in raw form (28-60 rows of folio
    numbers the model would be tempted to arithmetic on), every prose
    asset (Director's Message, Thank You), and any previous summary text.
    What goes in is computed scalars plus short structured records, so
    every figure the model can write is one this pipeline already
    computed and already prints elsewhere in the report.
    """
    from pipeline.risk_profile import compute_headline_equity_exposure_pct

    ps = ctx.portfolio_summary
    risk = ctx.risk_profile_result
    tax = ctx.tax_result

    top, bottom = _return_ranked_holdings(ctx.holdings)

    return {
        "client_name": ctx.client_name,
        "report_date": ctx.report_date.strftime("%d %b %Y"),

        "total_invested_rupees": _round_or_none(ps.total_invested, 0),
        "total_invested_display": format_inr(ps.total_invested),
        "current_value_rupees": _round_or_none(ps.current_value, 0),
        "current_value_display": format_inr(ps.current_value),
        "absolute_gain_rupees": _round_or_none(ps.absolute_gain, 0),
        "absolute_gain_display": format_inr(ps.absolute_gain),
        "absolute_gain_pct_display": _format_pct(ps.absolute_gain_pct),
        "value_weighted_cagr_pct_display": _format_pct(ps.portfolio_cagr_pct)
        if ps.portfolio_cagr_pct is not None else None,
        "cagr_method": (
            "value-weighted average of scheme-level CAGRs, not an XIRR"
        ),
        "number_of_schemes": ps.num_schemes,

        # Display string ONLY, at the report's own precision. Carrying the
        # raw 95.66 alongside it is what let the summary print a figure the
        # gauge and the allocation footnote both render as 95.7%.
        "equity_exposure_display": _equity_exposure_display(ctx.holdings),
        "risk_profile": risk.profile,
        "risk_profile_band": risk.band_definition,

        "top_holdings_by_return": top,
        "bottom_holdings_by_return": bottom,

        "transactions": [
            {
                "scheme": t.scheme,
                "action": t.action,
                "amount_rupees": _round_or_none(t.amount, 0),
                "amount_display": format_inr(t.amount) if t.amount is not None else None,
                "suggested_scheme": t.suggested_scheme,
            }
            for t in ctx.transaction_snapshot
        ],

        "total_computed_tax_rupees": _round_or_none(tax.summary.total_computed_tax, 0),
        "total_computed_tax_display": format_inr(tax.summary.total_computed_tax),
        "tax_zero_reason": _tax_zero_reason(tax),
        # The s.112A annual exemption LIMIT, not the gain that fell under
        # it. An earlier version passed equity_ltcg_exemption_applied,
        # whose name led the model to call Rs 48,000 "the LTCG exemption"
        # when Rs 48,000 was the exempted gain and the exemption is
        # Rs 1.25 lakh. A correctly named constant cannot be misread that
        # way, and it keeps a summary that mentions the threshold from
        # tripping validation on a figure the report genuinely uses.
        "ltcg_annual_exemption_limit_rupees": 125000,

        "earliest_deadline": _earliest_deadline(ctx.things_to_do),
        "things_to_do": [
            {
                "action": i.action,
                "scheme": i.scheme,
                "what_to_do": i.what_to_do,
                "deadline": i.deadline,
                "priority": i.priority,
            }
            for i in ctx.things_to_do
        ],
    }


# --------------------------------------------------------------------------
# 2. OUTPUT VALIDATION
# --------------------------------------------------------------------------

BANNED_PHRASES = (
    "target allocation",
    "drift",
    "rebalance to",
    "guaranteed",
    "will grow",
    "expected to return",
)

# Indian-format numeric token: optional rupee sign, digit groups with
# commas, optional decimals, optional scale word or percent sign.
_NUMBER_TOKEN = re.compile(
    r"(?:₹|Rs\.?\s*)?"
    r"(\d[\d,]*(?:\.\d+)?)"
    r"(?:\s*(crore|crores|cr|lakh|lakhs|lac)\b|\s*(%))?",
    re.IGNORECASE,
)

_SCALE = {
    "crore": 1e7, "crores": 1e7, "cr": 1e7,
    "lakh": 1e5, "lakhs": 1e5, "lac": 1e5,
}

# A fund name is a run of capitalised words containing "Fund" or "Scheme".
_FUND_NAME = re.compile(
    r"\b(?:[A-Z][\w&.\-]*\s+){1,6}(?:Fund|Scheme)\b(?:\s*\((?:G|D|IDCW)\))?"
)


@dataclass
class ValidationResult:
    ok: bool
    failures: list = field(default_factory=list)

    def reason(self) -> str:
        return "; ".join(self.failures)


def _collect_allowed_numbers(payload: Any, into: set) -> set:
    """Every numeric value anywhere in the input dict, at any depth."""
    if isinstance(payload, bool):
        return into
    if isinstance(payload, (int, float)):
        if not (isinstance(payload, float) and math.isnan(payload)):
            into.add(float(payload))
    elif isinstance(payload, dict):
        for value in payload.values():
            _collect_allowed_numbers(value, into)
    elif isinstance(payload, (list, tuple)):
        for value in payload:
            _collect_allowed_numbers(value, into)
    elif isinstance(payload, str):
        # Dates and deadlines are supplied as strings ("16 Aug 2026") but
        # the model may legitimately write the day or year as a numeral.
        for token in re.findall(r"\d+(?:\.\d+)?", payload):
            into.add(float(token))
    return into


def _structural_numbers(payload: dict) -> set:
    """Counts the model can legitimately state - "three transactions",
    "four items on the to-do list" - which are properties of the payload
    rather than values inside it."""
    counts = {0.0, 1.0}
    for key in ("transactions", "things_to_do", "top_holdings_by_return",
                "bottom_holdings_by_return"):
        counts.add(float(len(payload.get(key) or [])))
    return counts


def _extract_numbers(text: str) -> list:
    """Every numeric token in the response, normalised to a plain float.
    'Rs 1.25 crore' -> 12500000.0; '₹1,23,456' -> 123456.0; '12.5%' -> 12.5.

    Digits that are part of an identifier rather than a quantity - the 80
    in "Section 80C", the 112 in "s.112A" - are skipped. They are labels,
    not figures, and treating them as figures made the validator reject
    its own fallback text for quoting a section number.
    """
    found = []
    for match in _NUMBER_TOKEN.finditer(text):
        tail = text[match.end():match.end() + 1]
        if tail.isalnum():
            continue
        raw = match.group(1)
        suffix = (match.group(2) or "").lower()
        try:
            value = float(raw.replace(",", ""))
        except ValueError:  # pragma: no cover - regex guarantees digits
            continue
        if suffix in _SCALE:
            value *= _SCALE[suffix]
        found.append((match.group(0).strip(), value, len((raw.split(".") + [""])[1])))
    return found


def _matches_an_allowed_number(value: float, decimals: int, allowed: set) -> bool:
    """A written figure matches an input figure if they are equal, or
    equal once the input is rounded to the precision the model actually
    wrote. Rounding to the model's own stated precision is bounded and
    checkable - "12.5%" may stand for 12.53%, but no rounding of any
    input value produces a figure that was never in the input at all.

    Magnitudes match too: a portfolio that is down writes as "a loss of
    Rs 1,82,000", carrying the direction in words and the magnitude in
    figures, so +1,82,000 in the prose is the input's -1,82,000. This
    matters most on the fixtures where it is easiest to get tone wrong -
    a summary of a losing portfolio must be able to state the loss.
    """
    for candidate in allowed:
        for form in (candidate, abs(candidate)):
            if form == value:
                return True
            if round(form, decimals) == round(value, decimals):
                return True
            # Rupee figures are supplied unrounded but printed to the rupee.
            if decimals == 0 and round(form) == round(value):
                return True
    return False


def _known_fund_names(payload: dict) -> set:
    names = set()
    for key in ("top_holdings_by_return", "bottom_holdings_by_return"):
        for row in payload.get(key) or []:
            if row.get("scheme"):
                names.add(row["scheme"].lower())
    for row in payload.get("transactions") or []:
        for field_name in ("scheme", "suggested_scheme"):
            if row.get(field_name):
                names.add(row[field_name].lower())
    for row in payload.get("things_to_do") or []:
        if row.get("scheme"):
            names.add(row["scheme"].lower())
    return names


def validate_summary(text: str, payload: dict) -> ValidationResult:
    """Runs BEFORE any generated text is accepted. Returns every failure
    found, not just the first, so a retry can be told about all of them
    at once rather than fixing one and tripping the next."""
    failures = []

    lowered = text.lower()
    for phrase in BANNED_PHRASES:
        if phrase in lowered:
            failures.append(
                f"used the banned phrase {phrase!r} - the risk profile is inferred from the "
                f"current allocation, and this report makes no forward-return claims"
            )

    allowed = _collect_allowed_numbers(payload, set()) | _structural_numbers(payload)
    for written, value, decimals in _extract_numbers(text):
        if not _matches_an_allowed_number(value, decimals, allowed):
            failures.append(
                f"stated the figure {written!r}, which does not appear in the input data"
            )

    # Fund names are checked by MASKING every supplied name out of the
    # text first, then looking at what still parses as a fund name.
    # Matching candidate spans against the known set directly does not
    # work: a regex cannot reliably tell where a fund name starts or ends
    # in prose, so "Switch In Target Fund 1" yields the candidate "Switch
    # In Target Fund" - which neither contains nor is contained by the
    # real name "Target Fund 1". Masking sidesteps the boundary problem
    # entirely; whatever survives it was genuinely never supplied.
    known = _known_fund_names(payload)
    residual = text
    for name in sorted(known, key=len, reverse=True):
        residual = re.sub(re.escape(name), " ", residual, flags=re.IGNORECASE)
    for candidate in _FUND_NAME.findall(residual):
        stripped = candidate.strip()
        # A partial reference to a supplied name ("the Flexicap Fund" for
        # "HDFC Flexicap Fund") is a wording choice, not an invention.
        if any(stripped.lower() in k for k in known):
            continue
        failures.append(
            f"named the fund {stripped!r}, which is not in the input data"
        )

    # De-duplicated: the retry prompt should list each distinct problem
    # once, not repeat the same fund name for every sentence it appeared in.
    deduped = list(dict.fromkeys(failures))
    return ValidationResult(ok=not deduped, failures=deduped)


# --------------------------------------------------------------------------
# 3. DETERMINISTIC FALLBACK
# --------------------------------------------------------------------------

def _pct(value: Optional[float]) -> str:
    return "not available" if value is None else f"{value:.2f}%"


def render_fallback_summary(payload: dict) -> str:
    """Plain template built from the same input dict. Used when the API is
    unreachable, times out, or fails validation twice.

    This should read acceptably, not brilliantly. Its job is that the
    report always builds - including on a machine with no network - and
    that every figure in it is one the deterministic pipeline computed.
    """
    gain = payload.get("absolute_gain_rupees") or 0.0
    direction = "a gain of" if gain >= 0 else "a loss of"

    para1 = (
        f"As of {payload['report_date']}, your portfolio is invested across "
        f"{payload['number_of_schemes']} schemes. You have invested "
        f"{payload['total_invested_display']} and the current value is "
        f"{payload['current_value_display']}, which is {direction} "
        f"{format_inr(abs(gain))} ({payload['absolute_gain_pct_display']})."
    )
    if payload.get("value_weighted_cagr_pct_display"):
        para1 += (
            f" The value-weighted average of scheme-level CAGRs is "
            f"{payload['value_weighted_cagr_pct_display']} (not an XIRR)."
        )

    profile = payload.get("risk_profile")
    band = payload.get("risk_profile_band")
    equity = payload.get("equity_exposure_display")
    para2 = (
        f"Your equity exposure is {equity} of the portfolio."
        if equity else
        "Your equity exposure could not be computed from the uploaded file."
    )
    if profile:
        para2 += f" On that basis your computed risk profile is {profile}"
        para2 += f", a band covering {band}." if band else "."

    transactions = payload.get("transactions") or []
    if transactions:
        listed = "; ".join(
            f"{t['action']} {t['scheme']}"
            + (f" for {t['amount_display']}" if t.get("amount_display") else "")
            for t in transactions
        )
        para3 = f"The transactions proposed in this review are: {listed}."
    else:
        para3 = "No transactions are proposed in this review."

    para3 += f" The total computed tax on these transactions is {payload['total_computed_tax_display']}"
    reason = payload.get("tax_zero_reason")
    para3 += f" ({reason})." if reason else "."

    todos = payload.get("things_to_do") or []
    earliest = payload.get("earliest_deadline")
    if todos:
        listed = "; ".join(f"{i['action']} - {i['scheme']} (by {i['deadline']})" for i in todos)
        para4 = f"The follow-up items from this review are: {listed}."
        if earliest:
            para4 += (
                f" The earliest of these is {earliest['deadline']}: {earliest['what_to_do']} "
                f"({earliest['scheme']})."
            )
    else:
        para4 = "There are no follow-up items from this review."
    para4 += (
        " Your emergency fund and insurance cover were not available in the uploaded "
        "file and should be discussed at your next review with us."
    )

    return "\n\n".join([para1, para2, para3, para4])


# --------------------------------------------------------------------------
# 4. GENERATION
# --------------------------------------------------------------------------

@dataclass
class ClientSummary:
    """The Client Summary section's text, plus the human approval that
    gates it. `text` is editable - the RM review screen (prompt 7) writes
    back to it - and `approved` must be set True deliberately."""
    text: str
    approved: bool = False
    source: str = "fallback"           # "model" | "fallback"
    attempts: int = 0
    failure_log: list = field(default_factory=list)
    input_tokens: Optional[int] = None
    output_tokens: Optional[int] = None


def _user_prompt(payload: dict, previous_failure: Optional[str] = None) -> str:
    import json
    prompt = (
        "Here are the computed figures for this client's portfolio review. "
        "Write the Client Summary section from them.\n\n"
        + json.dumps(payload, indent=2, ensure_ascii=False, default=str)
    )
    if previous_failure:
        prompt += (
            "\n\nYour previous attempt was REJECTED by an automated check for the "
            f"following reason(s): {previous_failure}.\n"
            "Write the summary again, avoiding those problems. Every number you write "
            "must appear verbatim in the JSON above."
        )
    return prompt


def generate_client_summary(payload: dict, client=None) -> ClientSummary:
    """Generates the draft summary, validating before accepting it.

    Retries ONCE with the validation failure appended to the prompt. On a
    second failure - or any API error, timeout, or missing credential -
    falls back to the deterministic template and logs why. Never returns
    unvalidated model text.

    The returned summary is a DRAFT: `approved` is False, and
    build_report() will refuse to build until a human sets it True.
    """
    failure_log = []
    previous_failure = None

    for attempt in range(1, MAX_ATTEMPTS + 1):
        try:
            if client is None:
                import anthropic
                client = anthropic.Anthropic(
                    timeout=API_TIMEOUT_SECONDS, max_retries=API_MAX_RETRIES
                )
            response = client.messages.create(
                model=MODEL,
                max_tokens=MAX_TOKENS,
                temperature=TEMPERATURE,
                system=SYSTEM_PROMPT,
                messages=[{"role": "user", "content": _user_prompt(payload, previous_failure)}],
            )
        except Exception as exc:  # network, timeout, auth, rate limit - all fall back
            failure_log.append(f"attempt {attempt}: API call failed ({type(exc).__name__}: {exc})")
            break

        text = "".join(
            block.text for block in response.content if getattr(block, "type", None) == "text"
        ).strip()

        result = validate_summary(text, payload)
        if result.ok:
            return ClientSummary(
                text=text,
                approved=False,
                source="model",
                attempts=attempt,
                failure_log=failure_log,
                input_tokens=getattr(response.usage, "input_tokens", None),
                output_tokens=getattr(response.usage, "output_tokens", None),
            )

        failure_log.append(f"attempt {attempt}: validation failed - {result.reason()}")
        previous_failure = result.reason()

    return ClientSummary(
        text=render_fallback_summary(payload),
        approved=False,
        source="fallback",
        attempts=min(len(failure_log), MAX_ATTEMPTS),
        failure_log=failure_log,
    )
