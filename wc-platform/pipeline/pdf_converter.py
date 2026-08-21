"""
pipeline/pdf_converter.py

Converts the assembled report docx (pipeline/docx_builder.py) to PDF via a
headless LibreOffice subprocess, and resolves the Table of Contents page
numbers with a two-pass build:

  Pass 1: build_report() with toc_page_numbers=None -> placeholder "…" TOC.
          Convert to PDF.
  Scan:   Open the pass-1 PDF with PyMuPDF, walk every page from index 2
          onward (skipping the cover page and the TOC page itself), and
          record the page number where each SECTION_TITLES heading first
          appears - filtering spans to font-size >= 13pt so body text that
          happens to mention a section name (e.g. "see the Risk Profile
          section above") is never mistaken for the heading itself.
  Pass 2: build_report() again, this time with the detected page numbers
          injected into the TOC, and convert that final docx to PDF.

Known-bug fixes applied
------------------------
- H9: convert_docx_to_pdf() REQUIRES outdir to differ from the input
  docx's own directory. LibreOffice's --convert-to/--outdir is known to
  fail (or silently no-op) when they're the same directory - this is
  enforced with a hard check, not a suggestion.
- H4: detect_section_pages() only considers text spans with
  font-size >= 13pt as heading candidates. Without this filter, any body
  paragraph that happens to contain a section's exact title text gets
  wrongly picked up as "the" heading, corrupting the TOC page numbers.
"""

from __future__ import annotations

import shutil
import subprocess
from pathlib import Path
from typing import Callable, Optional, Union

import fitz  # PyMuPDF

from pipeline.docx_builder import SECTION_TITLES, ReportContext, build_report

MIN_HEADING_FONT_SIZE = 13.0
HEADING_SCAN_START_PAGE_INDEX = 2  # 0-based: skip page 1 (cover) and page 2 (TOC)


# --------------------------------------------------------------------------
# LibreOffice conversion
# --------------------------------------------------------------------------

def _find_soffice() -> str:
    candidates = [
        "soffice",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        "/usr/bin/soffice",
        "/usr/local/bin/soffice",
        "/opt/homebrew/bin/soffice",
    ]
    for c in candidates:
        if shutil.which(c) or Path(c).exists():
            return c
    raise FileNotFoundError(
        "soffice (LibreOffice) binary not found. Install LibreOffice "
        "(`brew install --cask libreoffice` on macOS, or `libreoffice` via "
        "packages.txt on Streamlit Cloud) or pass soffice_bin explicitly."
    )


def convert_docx_to_pdf(
    docx_path: Union[str, Path],
    outdir: Union[str, Path],
    soffice_bin: Optional[str] = None,
    timeout: int = 180,
) -> Path:
    """Converts a .docx to .pdf via headless LibreOffice.

    outdir MUST be a different directory from docx_path's parent - bug H9:
    same-directory conversion is known to fail. This is enforced, not
    merely documented.
    """
    docx_path = Path(docx_path)
    outdir = Path(outdir)

    if outdir.resolve() == docx_path.parent.resolve():
        raise ValueError(
            "convert_docx_to_pdf: outdir must differ from the input docx's own "
            "directory - same-directory LibreOffice conversion is a known failure "
            "mode (bug H9). Pass a separate output directory."
        )

    outdir.mkdir(parents=True, exist_ok=True)
    soffice_bin = soffice_bin or _find_soffice()

    cmd = [
        soffice_bin, "--headless", "--norestore",
        "--convert-to", "pdf", "--outdir", str(outdir), str(docx_path),
    ]
    result = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout)

    pdf_path = outdir / (docx_path.stem + ".pdf")
    if result.returncode != 0 or not pdf_path.exists():
        raise RuntimeError(
            f"LibreOffice conversion failed (exit code {result.returncode}).\n"
            f"stdout: {result.stdout}\nstderr: {result.stderr}"
        )
    return pdf_path


# --------------------------------------------------------------------------
# TOC heading detection (bug H4 fix: font-size filter)
# --------------------------------------------------------------------------

def detect_section_pages(
    pdf_path: Union[str, Path],
    section_titles: list = SECTION_TITLES,
    start_page_index: int = HEADING_SCAN_START_PAGE_INDEX,
    min_font_size: float = MIN_HEADING_FONT_SIZE,
) -> dict:
    """Scans a PDF for each section title, returning {title: page_number
    (1-based)}. Only text spans with font-size >= min_font_size are
    considered - this is the bug H4 fix. A title not found anywhere is
    simply absent from the returned dict (build_report()'s TOC then falls
    back to '…' for that entry rather than a guessed number)."""
    pdf_path = Path(pdf_path)
    doc = fitz.open(str(pdf_path))
    found: dict = {}
    try:
        remaining = set(section_titles)
        for page_index in range(start_page_index, len(doc)):
            if not remaining:
                break
            page = doc[page_index]
            text_dict = page.get_text("dict")
            for block in text_dict.get("blocks", []):
                for line in block.get("lines", []):
                    for span in line.get("spans", []):
                        if span["size"] < min_font_size:
                            continue  # bug H4: body text is excluded purely by size
                        text = span["text"].strip()
                        if text in remaining:
                            found[text] = page_index + 1  # 1-based page number
                            remaining.discard(text)
    finally:
        doc.close()
    return found


# --------------------------------------------------------------------------
# Two-pass orchestration
# --------------------------------------------------------------------------

class PdfUnavailable(RuntimeError):
    """LibreOffice is absent, so no PDF could be rendered. Carries the
    .docx that WAS built successfully, because losing it too would be a
    second failure caused by the first."""

    def __init__(self, message: str, docx_path: Path):
        super().__init__(message)
        self.docx_path = docx_path


def soffice_available(soffice_bin: Optional[str] = None) -> bool:
    """Whether a PDF can be rendered at all in this environment.

    Streamlit Community Cloud has no LibreOffice unless packages.txt
    installs it, and that install is large enough to fail. The app asks
    this up front so it can tell the RM what they are going to get BEFORE
    they click Generate, rather than after.
    """
    try:
        _find_soffice() if soffice_bin is None else soffice_bin
        if soffice_bin is not None:
            return bool(shutil.which(soffice_bin) or Path(soffice_bin).exists())
        return True
    except FileNotFoundError:
        return False


def build_report_deliverables(
    ctx: ReportContext,
    workdir: Union[str, Path],
    soffice_bin: Optional[str] = None,
    build_fn: Callable = build_report,
    require_pdf: bool = True,
) -> tuple:
    """Runs the two-pass TOC-resolution build and returns
    (final_pdf_path, final_docx_path, detected_page_numbers).

    BOTH returned files are user-facing deliverables: the PDF is the
    primary download, the .docx is the secondary option for manual
    editing. They are not alternatives generated by different code paths -
    the PDF is rendered FROM the returned docx, so any feature present in
    one is present in the other by construction. That's deliberate: it
    removes the possibility of a feature working in one output and
    silently degrading in the other.

    The returned docx is the PASS-2 file (real TOC page numbers), never
    the pass-1 placeholder file - handing a user a docx whose contents
    page reads "…" would be a silent degradation of exactly the kind this
    function exists to prevent.
    """
    workdir = Path(workdir)
    docx_dir = workdir / "docx"
    pdf_dir = workdir / "pdf"
    docx_dir.mkdir(parents=True, exist_ok=True)
    pdf_dir.mkdir(parents=True, exist_ok=True)

    # One timestamp for the whole run, so the delivered docx and the PDF
    # rendered from it are stamped identically (and both passes agree).
    if ctx.generated_at is None:
        from datetime import datetime
        ctx.generated_at = datetime.now()

    # --- No LibreOffice: ship the docx rather than nothing ------------
    #
    # The two-pass TOC resolver needs a rendered PDF to read page numbers
    # off, so without LibreOffice there is no way to fill the contents
    # page - and no way to produce a PDF at all. What there IS is a
    # complete, correct .docx, built before conversion is even attempted.
    # Raising past it would throw away a working deliverable because a
    # secondary one could not be made. The TOC keeps its placeholder dots
    # and the caller is told exactly that.
    if not soffice_available(soffice_bin):
        fallback_docx = docx_dir / "report_final.docx"
        build_fn(ctx, fallback_docx, toc_page_numbers=None)
        if require_pdf:
            raise PdfUnavailable(
                "LibreOffice (soffice) is not available in this environment, so the PDF could "
                "not be rendered and the contents page could not be resolved to real page "
                "numbers. The editable .docx was built successfully and is complete apart from "
                "those page numbers.",
                docx_path=fallback_docx,
            )
        return None, fallback_docx, {}

    # --- Pass 1: placeholder TOC ---
    pass1_docx = docx_dir / "report_pass1.docx"
    build_fn(ctx, pass1_docx, toc_page_numbers=None)
    pass1_pdf_dir = pdf_dir / "pass1"  # H9: separate from docx_dir
    pass1_pdf = convert_docx_to_pdf(pass1_docx, pass1_pdf_dir, soffice_bin=soffice_bin)

    # --- Scan for heading page numbers ---
    detected_pages = detect_section_pages(pass1_pdf)

    # --- Pass 2: real TOC page numbers. This docx IS the deliverable. ---
    pass2_docx = docx_dir / "report_final.docx"
    build_fn(ctx, pass2_docx, toc_page_numbers=detected_pages)
    pass2_pdf_dir = pdf_dir / "pass2"  # H9: separate from docx_dir, and from pass1's outdir
    pass2_pdf = convert_docx_to_pdf(pass2_docx, pass2_pdf_dir, soffice_bin=soffice_bin)

    return pass2_pdf, pass2_docx, detected_pages


def build_report_pdf_two_pass(ctx, workdir, soffice_bin=None, build_fn=build_report) -> tuple:
    """Backwards-compatible wrapper returning only (pdf, detected_pages).
    Prefer build_report_deliverables(), which also returns the docx - both
    files are deliverables now, not just the PDF."""
    pdf_path, _docx_path, detected = build_report_deliverables(
        ctx, workdir, soffice_bin=soffice_bin, build_fn=build_fn,
    )
    return pdf_path, detected


# --------------------------------------------------------------------------
# Self-test / end-to-end dummy run
# --------------------------------------------------------------------------

def _build_dummy_context(assets_dir: Path, mindmap_png_path: Path) -> ReportContext:
    from datetime import date

    from pipeline.docx_builder import (
        FirmInfo, PerformanceRow, PortfolioSummary, RMInfo,
        ThingsToDoRow, TransactionSnapshotRow,
    )
    from pipeline.chart_gen import format_inr
    from pipeline.parser import Holding, compute_value_weighted_cagr
    from pipeline.risk_profile import RiskHolding, classify_scheme_market_cap, compute_risk_profile
    from pipeline.tax_calc import (
        allocate_tax_across_transactions, build_tax_holdings_from_transactions, compute_portfolio_tax,
    )

    # 27 holdings (23 Equity / 2 Hybrid / 1 Debt / 1 Other) - a realistic
    # ~25-30-holding portfolio, deliberately mixing short and long scheme
    # names (including a couple as long as the benchmark's "Canara Robeco
    # Large and Mid Cap Fund Reg (G)"), to genuinely test the Holdings
    # Statement density fix rather than just re-checking the original
    # 5-holding smoke test.
    _equity_specs = [
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
    holdings = []
    for scheme, folio, units, purchase, current in _equity_specs:
        gain = current - purchase
        holdings.append(Holding(
            member="Rahul Sharma", pan="ABCDE1234F", category="Equity",
            scheme=scheme, folio=folio, balance_units=units,
            purchase_value=purchase, current_value=current, gain=gain,
            absolute_return_pct=round(gain / purchase * 100, 2),
            cagr_pct=round(gain / purchase * 100 / 3, 2),  # rough illustrative CAGR for dummy data
        ))

    holdings.append(Holding(
        member="Rahul Sharma", pan="ABCDE1234F", category="Hybrid",
        scheme="ICICI Pru Equity & Debt Fund", folio="334455667788", balance_units=2210.5,
        purchase_value=300000, current_value=356000, gain=56000,
        absolute_return_pct=18.7, cagr_pct=13.0))
    holdings.append(Holding(
        member="Rahul Sharma", pan="ABCDE1234F", category="Hybrid",
        scheme="HSBC Aggressive Hybrid Fund Reg (G)", folio="345988815", balance_units=1980.221,
        purchase_value=180000, current_value=214300, gain=34300,
        absolute_return_pct=19.06, cagr_pct=12.4))
    holdings.append(Holding(
        member="Rahul Sharma", pan="ABCDE1234F", category="Debt",
        scheme="ICICI Pru Short Term Fund", folio="556677889900", balance_units=15320.0,
        purchase_value=200000, current_value=215000, gain=15000,
        absolute_return_pct=7.5, cagr_pct=6.2))
    holdings.append(Holding(
        member="Rahul Sharma", pan="ABCDE1234F", category="Other",
        scheme="SBI Gold Fund Direct (G)", folio="144075232", balance_units=2851.662,
        purchase_value=90000, current_value=118400, gain=28400,
        absolute_return_pct=31.56, cagr_pct=18.9))
    holdings.append(Holding(
        member="Rahul Sharma", pan="ABCDE1234F", category="Equity",
        scheme="Mirae Asset ELSS Tax Saver Fund Reg (G)", folio="775149452599", balance_units=3210.445,
        purchase_value=110000, current_value=142300, gain=32300,
        absolute_return_pct=29.36, cagr_pct=15.2))

    def _infer_fund_type(h) -> str:
        # Test-data-only heuristic for risk_profile.py's EQUITY/HYBRID/DEBT
        # look-through WEIGHTING only (i.e. "is this 100% equity, 75%
        # equity, 0% equity") - NOT the market-cap classification, which
        # is now done separately from the actual scheme name via
        # risk_profile.classify_scheme_market_cap(). "Large Cap Fund" here
        # is just a stand-in that reliably matches EQUITY_SUBCATEGORY_
        # RULES as "100% equity" for every Equity-category holding.
        if h.category == "Equity":
            return "Large Cap Fund"
        if h.category == "Hybrid":
            return "Aggressive Hybrid Fund"
        if h.category == "Debt":
            return "Debt Fund"
        if h.category == "Other":
            return "Gold ETF" if "gold" in (h.scheme or "").lower() else "Other Fund"
        return "Other Fund"

    risk_holdings = [RiskHolding(h.scheme, _infer_fund_type(h), h.current_value) for h in holdings]
    risk_result = compute_risk_profile(risk_holdings)

    as_of = date(2026, 8, 16)

    # Derived directly from `holdings` (Other -> Gold/SGB) - NO extra
    # hardcoded categories. asset_allocation must sum to exactly the
    # holdings grand total (docx_builder._validate_report_context asserts
    # this at build time) - it previously carried two invented slivers
    # ("Global Equity/International" ₹60,000, "Other/Liquid/Cash"
    # ₹15,000) with no matching holding at all, which is exactly the bug
    # fix #1 is about.
    _category_to_pie_bucket = {"Equity": "Equity", "Hybrid": "Hybrid", "Debt": "Debt", "Other": "Gold/SGB"}
    asset_allocation: dict = {}
    for h in holdings:
        bucket = _category_to_pie_bucket.get(h.category, "Other/Liquid/Cash")
        asset_allocation[bucket] = asset_allocation.get(bucket, 0) + (h.current_value or 0)

    _total_invested = sum(h.purchase_value for h in holdings if h.purchase_value is not None)
    _total_current = sum(h.current_value for h in holdings if h.current_value is not None)
    portfolio_summary = PortfolioSummary(
        total_invested=_total_invested,
        current_value=_total_current,
        absolute_gain=_total_current - _total_invested,
        absolute_gain_pct=(_total_current - _total_invested) / _total_invested * 100,
        # Actually COMPUTED from the holdings (value-weighted average of
        # each holding's own CAGR) - was previously a hardcoded 12.8 that
        # never moved even when the holdings list and total invested did.
        portfolio_cagr_pct=compute_value_weighted_cagr(holdings),
        monthly_sip=25000,
        # UNIQUE schemes, not folio rows - "Canara Robeco Large and Mid Cap
        # Fund Reg (G)" is held under two separate folios and must count
        # once. The Holdings table still renders one row per folio.
        num_schemes=len({h.scheme for h in holdings if h.scheme}),
    )

    # --- Transactions are built in two phases, because a Switch In's
    # amount depends on the tax allocated to its paired Switch Out, and
    # that tax can only be computed once ALL proposed transactions are
    # known (the Rs 1.25L s.112A exemption is annual and shared - it
    # cannot be applied per transaction in isolation). ---
    #
    # Phase 1: every row except the Switch In amounts, which are pending.
    _hdfc_switch_out_amount = 298000        # current value of the 980.234 units
    _hdfc_switch_out_purchase = 250000
    _hdfc_switch_out_purchase_date = date(2024, 1, 15)

    transaction_snapshot = [
        # Amount = CURRENT value of the 980.234 units being switched out
        # (₹2,98,000, matching the HDFC Flexicap holding's current_value
        # above) - NOT the ₹2,50,000 purchase value, which is kept
        # separately in purchase_amount. purchase_date is required here
        # (not on Holding) so the Tax Analysis section can classify this
        # as LTCG/STCG - see build_tax_holdings_from_transactions below.
        TransactionSnapshotRow("HDFC Flexicap Fund", "Switch Out", _hdfc_switch_out_amount,
                                balance_units=980.234,
                                purchase_amount=_hdfc_switch_out_purchase,
                                suggested_scheme="Parag Parikh Flexicap Fund",
                                purchase_date=_hdfc_switch_out_purchase_date),
        # Amount filled in during phase 2 below, once the shared-exemption
        # tax allocation is known.
        TransactionSnapshotRow("Parag Parikh Flexicap Fund", "Switch In", None),
        # Amount already equalled current_value (₹71,500) here - no change needed.
        TransactionSnapshotRow("L&T Emerging Businesses Fund", "Redeem", 71500, balance_units=560.912,
                                purchase_amount=80000, purchase_date=date(2026, 3, 1)),
        TransactionSnapshotRow("Kotak Multicap Fund", "Reinvest", 71500),
        # SIP Stop row deliberately has NO balance_units / purchase_amount - must render blank, not inferred.
        # SIP Stop is not a redemption event either - no purchase_date, and
        # it's excluded from tax computation by action type regardless.
        TransactionSnapshotRow("Axis Bluechip Fund", "SIP Stop", 10000, suggested_scheme="Mirae Asset Large Cap Fund"),
        TransactionSnapshotRow("Mirae Asset Large Cap Fund", "SIP Start", 10000),
    ]

    # Phase 2: compute the total tax across ALL proposed transactions once,
    # apportion it pro-rata, and net each Switch Out's share out of its
    # paired Switch In. Pairing is by the Switch Out's suggested_scheme.
    _pending_tax_holdings, _ = build_tax_holdings_from_transactions(
        transaction_snapshot, holdings, as_of=as_of,
    )
    _tax_allocation = allocate_tax_across_transactions(_pending_tax_holdings, as_of=as_of)

    _switch_in_by_scheme = {t.scheme: t for t in transaction_snapshot if t.action == "Switch In"}
    for t in transaction_snapshot:
        if t.action != "Switch Out":
            continue
        allocated = _tax_allocation.get(t.scheme, 0.0)
        t.switch_deduction = allocated
        t.switch_deduction_note = "capital gains tax (this switch's share of the total)"
        paired = _switch_in_by_scheme.get(t.suggested_scheme)
        if paired is not None:
            paired.amount = t.amount - allocated

    # Capital-gains tax is computed ONLY on the Switch Out / Redeem rows
    # above - the SAME list that renders in the Transaction Snapshot table,
    # not a separately-maintained holdings list. Untransacted holdings'
    # unrealised gains are never taxed in this section.
    tax_holdings, tax_holdings_warnings = build_tax_holdings_from_transactions(
        transaction_snapshot, holdings, as_of=as_of,
    )
    for w in tax_holdings_warnings:
        print(f"[tax] warning: {w}")

    tax_result = compute_portfolio_tax(tax_holdings, as_of=as_of)

    performance_rows = [
        PerformanceRow("HDFC Flexicap Fund", "out", {
            "1Y": 14.28, "2Y": 11.2, "3Y": 13.5, "5Y": 12.1, "7Y": 11.8, "10Y": 11.55,
            "Since Launch": 13.0, "CY": 9.5, "CY-1": 14.0, "CY-2": 8.0, "CY-3": -5.2, "CY-4": 5.0,
        }),
        PerformanceRow("L&T Emerging Businesses Fund", "out", {
            "1Y": -3.4, "2Y": 9.8, "3Y": 16.2, "5Y": 15.0, "7Y": "N/A", "10Y": "N/A",
            "Since Launch": 14.5, "CY": -3.4, "CY-1": 22.1, "CY-2": 4.2, "CY-3": -8.1, "CY-4": 12.0,
        }),
        PerformanceRow("Parag Parikh Flexicap Fund", "in", {
            "1Y": 18.9, "2Y": 16.4, "3Y": 19.1, "5Y": 18.0, "7Y": 17.2, "10Y": "N/A",
            "Since Launch": 19.5, "CY": 12.1, "CY-1": 21.0, "CY-2": 6.5, "CY-3": 2.3, "CY-4": 15.4,
        }),
        PerformanceRow("Kotak Multicap Fund", "in", {
            "1Y": 12.6, "2Y": "N/A", "3Y": "N/A", "5Y": "N/A", "7Y": "N/A", "10Y": "N/A",
            "Since Launch": 13.1, "CY": 8.0, "CY-1": 10.5, "CY-2": "N/A", "CY-3": "N/A", "CY-4": "N/A",
        }),
    ]

    # NOTE: there is deliberately NO "rebalance back to target allocation"
    # action here. WC infers the risk profile FROM the current allocation,
    # so there is no target band to drift from - such an action would
    # describe a concept this model doesn't have.
    #
    # `priority` drives the rendered order (and the "#" column is renumbered
    # from it), so the Emergency Fund row - whose own section text calls it
    # "Priority 1" - actually appears first.
    things_to_do = [
        ThingsToDoRow(0, "Review", "Emergency Fund",
                      "Discuss liquid-fund buffer with client - no data on file this cycle",
                      "15 Sep 2026", priority=1),
        ThingsToDoRow(0, "Confirm", "HDFC Flexicap Fund",
                      "Get client sign-off on switch to Parag Parikh Flexicap Fund",
                      "30 Aug 2026", priority=2),
        ThingsToDoRow(0, "Update", "KYC / Nominee",
                      "Confirm nominee details are current across all folios",
                      "31 Oct 2026", priority=4),
    ]

    return ReportContext(
        client_name="Rahul Sharma",
        client_salutation="Mr.",
        report_date=date(2026, 8, 16),
        firm=FirmInfo(),
        rm=RMInfo(name="Relationship Manager", email="rm@wealthcareindia.com", phone="+91-98100-00000"),
        logo_path=assets_dir / "logo_2.jpg",
        portfolio_summary=portfolio_summary,
        asset_allocation=asset_allocation,
        equity_sub_allocation=risk_result.equity_sub_allocation,
        risk_profile_result=risk_result,
        holdings=holdings,
        mindmap_path=mindmap_png_path,
        transaction_snapshot=transaction_snapshot,
        performance_rows=performance_rows,
        tax_result=tax_result,
        emergency_fund_insurance=None,  # deliberately absent - exercises the fallback text
        things_to_do=things_to_do,
        director_message_path=assets_dir / "director_message.docx",
        thank_you_message_path=assets_dir / "thank_you_message.docx",
    )


def _run_self_test() -> None:
    from pipeline.mindmap import build_mindmap_recommendations_from_transactions, generate_mindmap

    print("=== pipeline/pdf_converter.py end-to-end self-test ===\n")

    project_root = Path(__file__).parent.parent
    assets_dir = project_root / "assets"
    output_dir = project_root / "_test_output"
    output_dir.mkdir(exist_ok=True)

    mindmap_path = output_dir / "mindmap.png"

    # Build the context FIRST (with a placeholder mindmap path), so the
    # mind map can be generated from ctx.transaction_snapshot - the exact
    # same list the Transaction Snapshot table renders. The mind map used
    # to be built from its own separately-maintained recommendation list,
    # which is how it ended up showing a different amount for the same
    # transaction than the Snapshot did.
    ctx = _build_dummy_context(assets_dir, mindmap_path)

    mindmap_recs = build_mindmap_recommendations_from_transactions(ctx.transaction_snapshot)
    mindmap_result = generate_mindmap(mindmap_recs, client_name=ctx.client_name, output_path=mindmap_path)
    print(f"Mind map rendered: {mindmap_result.png_path} (warnings: {mindmap_result.warnings})")
    print(f"Mind map amounts (from transaction_snapshot): "
          f"{[(r.scheme, r.action, r.amount) for r in mindmap_recs]}")

    print("\nRunning two-pass build...")
    final_pdf, final_docx, detected_pages = build_report_deliverables(ctx, workdir=output_dir)

    print(f"\nDetected TOC page numbers (from pass-1 PDF scan):")
    for title in SECTION_TITLES:
        print(f"  {title:<36} -> page {detected_pages.get(title, '(not found)')}")

    # BOTH files are deliverables - the PDF is the primary download, the
    # docx the secondary editable copy. Both are written on every run.
    stable_docx = output_dir / "Portfolio_Review_Report_Rahul_Sharma.docx"
    stable_pdf = output_dir / "Portfolio_Review_Report_Rahul_Sharma.pdf"
    stable_docx.write_bytes(final_docx.read_bytes())
    stable_pdf.write_bytes(final_pdf.read_bytes())

    print(f"\nDeliverable 1 (primary)   PDF:  {stable_pdf} ({stable_pdf.stat().st_size:,} bytes)")
    print(f"Deliverable 2 (editable)  DOCX: {stable_docx} ({stable_docx.stat().st_size:,} bytes)")
    print(f"Generation timestamp stamped in both footers: {ctx.generated_at:%d %b %Y at %H:%M}")

    doc = fitz.open(str(stable_pdf))
    print(f"PDF page count: {len(doc)}")
    doc.close()

    print("\nSelf-test completed successfully.")


if __name__ == "__main__":
    _run_self_test()
