#!/usr/bin/env python
"""
scripts/run_real_summary.py

Makes the two real Anthropic API calls for the Client Summary - one
against A_golden, one against G_all_loss - and prints each generated
summary with its token counts.

    export ANTHROPIC_API_KEY=sk-ant-...
    ./.venv/bin/python scripts/run_real_summary.py

G_all_loss is not optional padding: a portfolio down 20% is where tone
and honesty matter most, and where a model is most tempted to reassure.
Reading that output is the point of running this.

Everything printed here is a DRAFT. `approved` is False on both, and
docx_builder.build_report() refuses to build until a human sets it True.
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from pipeline.summary_client import (  # noqa: E402
    MODEL,
    MAX_TOKENS,
    TEMPERATURE,
    build_summary_input,
    generate_client_summary,
    validate_summary,
)
from tests.fixtures import ALL_FIXTURE_BUILDERS  # noqa: E402

FIXTURES = ("A_golden", "G_all_loss")


def main() -> int:
    import tempfile

    print(f"model={MODEL}  max_tokens={MAX_TOKENS}  temperature={TEMPERATURE}\n")
    totals = {"in": 0, "out": 0}
    exit_code = 0

    with tempfile.TemporaryDirectory() as tmp:
        for name in FIXTURES:
            ctx = ALL_FIXTURE_BUILDERS[name](mindmap_path=Path(tmp) / f"{name}.png").ctx
            payload = build_summary_input(ctx)
            summary = generate_client_summary(payload)

            print("=" * 78)
            print(f"{name}   source={summary.source}   attempts={summary.attempts}")
            print(f"tokens: input={summary.input_tokens}  output={summary.output_tokens}")
            if summary.failure_log:
                for entry in summary.failure_log:
                    print(f"  ! {entry}")
            print("=" * 78)
            print(summary.text)
            print()

            if summary.source == "fallback":
                exit_code = 1  # the API call did not land - see the failure log above
            totals["in"] += summary.input_tokens or 0
            totals["out"] += summary.output_tokens or 0

            result = validate_summary(summary.text, payload)
            assert result.ok, f"[{name}] returned text failed validation: {result.reason()}"
            assert summary.approved is False, "generated text must arrive unapproved"

    print(f"TOTAL TOKENS  input={totals['in']}  output={totals['out']}")
    return exit_code


if __name__ == "__main__":
    raise SystemExit(main())
