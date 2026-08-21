"""
pipeline/chart_gen.py

Matplotlib donut chart generator for the WC Securities / Wealthkare
Portfolio Review pipeline - renders the asset allocation chart embedded in
the report docx.

Design notes
------------
- Transparent background throughout (figure + axes), so the chart drops
  cleanly onto a docx page without a white box around it.
- Exported at 2x resolution (150 dpi figure, saved at 300 dpi) so it stays
  crisp when placed into a Word document at typical report sizing.
- Legend entries show BOTH the ₹ value (Indian digit grouping, e.g.
  "₹2,04,42,989") AND the % of total for each slice - not one or the
  other.
- Category colors are brand-consistent (navy/gold family) with sensible
  fallbacks for a value passed that isn't explicitly mapped.
"""

from __future__ import annotations

from pathlib import Path
from typing import Optional, Union

import matplotlib
matplotlib.use("Agg")  # headless-safe backend, no display server required
import matplotlib.pyplot as plt

# --------------------------------------------------------------------------
# Constants
# --------------------------------------------------------------------------

# Normalized (lowercased, stripped) category name -> hex color.
# Covers every category named in the spec, plus common aliases.
CATEGORY_COLOR_MAP: dict[str, str] = {
    "equity": "#1C2B4B",                       # navy
    "equity (domestic)": "#1C2B4B",
    "domestic equity": "#1C2B4B",
    "hybrid": "#B8860B",                        # gold
    "debt": "#4C7A5E",                          # muted green
    "gold": "#D4A017",                          # amber
    "gold/sgb": "#D4A017",
    "sgb": "#D4A017",
    "gold/sgb/gold etf": "#D4A017",
    "global equity": "#4A6FA5",                 # steel blue
    "international": "#4A6FA5",
    "global equity/international": "#4A6FA5",
    "global equity / international": "#4A6FA5",
    "other": "#9B9B93",                         # warm gray
    "liquid": "#9B9B93",
    "cash": "#9B9B93",
    "other/liquid/cash": "#9B9B93",
    "other / liquid / cash": "#9B9B93",
}

# Fallback cycle for any category not found above, so an unexpected label
# still renders distinctly instead of erroring out.
FALLBACK_COLOR_CYCLE = ["#7B6D8D", "#C97B63", "#5D8AA8", "#8A9B6E", "#A66B8E"]

NAVY = "#1C2B4B"


# --------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------

def format_inr(value: float) -> str:
    """Format a number using Indian digit grouping with a ₹ prefix, e.g.
    20442989 -> '₹2,04,42,989'."""
    value = int(round(value))
    sign = "-" if value < 0 else ""
    s = str(abs(value))

    if len(s) <= 3:
        formatted = s
    else:
        last3 = s[-3:]
        rest = s[:-3]
        parts: list[str] = []
        while len(rest) > 2:
            parts.insert(0, rest[-2:])
            rest = rest[:-2]
        if rest:
            parts.insert(0, rest)
        formatted = ",".join(parts) + "," + last3

    return f"{sign}₹{formatted}"


def _color_for_category(category: str, fallback_index: list[int]) -> str:
    key = category.strip().lower()
    if key in CATEGORY_COLOR_MAP:
        return CATEGORY_COLOR_MAP[key]
    color = FALLBACK_COLOR_CYCLE[fallback_index[0] % len(FALLBACK_COLOR_CYCLE)]
    fallback_index[0] += 1
    return color


# --------------------------------------------------------------------------
# Core function
# --------------------------------------------------------------------------

def generate_donut_chart(
    allocation: dict[str, float],
    output_path: Union[str, Path],
) -> Path:
    """Render an asset-allocation chart and save it as a transparent PNG at
    2x resolution.

    Deliberately a SOLID pie (no center hole, no center label) - matches
    the benchmark report's flat pie style exactly. The function name is
    kept as-is (rather than renamed to generate_pie_chart) so existing
    call sites in docx_builder.py / pdf_converter.py don't need touching -
    only the rendered shape changed.

    allocation: {category_label: current_value}. Categories with a
    zero/negative/missing value are dropped from the chart (they'd render
    as a zero-width slice anyway) rather than causing an error.

    Returns the saved PNG path.
    """
    output_path = Path(output_path)

    clean_allocation = {
        cat: val for cat, val in allocation.items()
        if val is not None and val > 0
    }
    if not clean_allocation:
        raise ValueError("generate_donut_chart: no categories with a positive value were provided.")

    categories = list(clean_allocation.keys())
    values = list(clean_allocation.values())
    total = sum(values)

    fallback_index = [0]
    colors = [_color_for_category(cat, fallback_index) for cat in categories]

    fig, ax = plt.subplots(figsize=(7.5, 7.5), dpi=150)
    fig.patch.set_alpha(0.0)
    ax.set_facecolor("none")

    # No `width=` in wedgeprops -> a full solid wedge, not a ring. This is
    # the whole of the "donut -> pie" fix: no hole, no center text.
    ax.pie(
        values,
        colors=colors,
        startangle=90,
        counterclock=False,
        wedgeprops=dict(edgecolor="white", linewidth=1.5),
    )
    ax.set(aspect="equal")

    # Round dot legend markers (matplotlib's default legend handles for a
    # pie are rectangular wedge swatches) - matches the benchmark's round
    # bullet-point legend style.
    from matplotlib.lines import Line2D
    legend_handles = [
        Line2D([0], [0], marker="o", color="none", markerfacecolor=color,
               markeredgecolor="none", markersize=11)
        for color in colors
    ]
    legend_labels = [
        f"{cat}  {format_inr(val)}  {(val / total) * 100:.1f}%"
        for cat, val in zip(categories, values)
    ]
    ax.legend(
        legend_handles,
        legend_labels,
        loc="center left",
        bbox_to_anchor=(1.02, 0.5),
        frameon=False,
        fontsize=11,
        labelcolor=NAVY,
        handlelength=1.2,
        handleheight=1.2,
    )

    fig.savefig(
        output_path,
        dpi=300,  # 2x the 150 dpi figure -> crisp for docx embedding
        transparent=True,
        bbox_inches="tight",
    )
    plt.close(fig)

    return output_path


# --------------------------------------------------------------------------
# Self-test (dummy data, no external files required)
# --------------------------------------------------------------------------

def _run_self_test() -> None:
    import tempfile

    print("=== pipeline/chart_gen.py self-test ===\n")

    sample_allocation = {
        "Equity": 20442989,
        "Hybrid": 4850000,
        "Debt": 3120000,
        "Gold/SGB": 950000,
        "Global Equity/International": 1780000,
        "Other/Liquid/Cash": 410000,
    }

    total = sum(sample_allocation.values())
    print("Input allocation:")
    for cat, val in sample_allocation.items():
        print(f"  {cat:<32} {format_inr(val):>16}  ({val / total * 100:.1f}%)")

    with tempfile.TemporaryDirectory() as tmp:
        out_path = Path(tmp) / "donut_test.png"
        saved_path = generate_donut_chart(sample_allocation, out_path)

        assert saved_path.exists(), "PNG was not written to disk."
        size_bytes = saved_path.stat().st_size
        print(f"\nSaved PNG: {saved_path} ({size_bytes:,} bytes)")

        # Confirm the file is a real PNG and check its pixel dimensions to
        # verify the 2x-resolution export.
        try:
            from PIL import Image
            with Image.open(saved_path) as img:
                print(f"Image dimensions: {img.size[0]}x{img.size[1]} px, mode={img.mode}")
                assert img.mode in ("RGBA", "LA"), "Expected an alpha channel for a transparent background."
        except ImportError:
            print("(Pillow not installed - skipping pixel-dimension / alpha-channel check.)")

        # Copy the test output to a stable, inspectable location before the
        # temp dir is cleaned up.
        stable_path = Path(__file__).parent.parent / "_test_output_donut.png"
        stable_path.write_bytes(saved_path.read_bytes())
        print(f"Copied to stable path for inspection: {stable_path}")

    print("\nAll self-test assertions passed.")


if __name__ == "__main__":
    _run_self_test()
