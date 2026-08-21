"""
tests/conftest.py

Shared pytest wiring for the regression suite.

Fixtures are built once per session (they're pure data derivation, and
the `large` one builds 60 holdings) and handed to the invariant tests
parameterized by name, so a failure message names the dataset it came
from.
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

# Make `pipeline` importable when pytest is invoked from anywhere.
PROJECT_ROOT = Path(__file__).parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from tests.fixtures import ALL_FIXTURE_BUILDERS  # noqa: E402

FIXTURE_NAMES = list(ALL_FIXTURE_BUILDERS)


@pytest.fixture(scope="session")
def built_fixtures() -> dict:
    """All 7 datasets, built once. Each build calls the production
    derivation functions, so a crash here is itself a real failure."""
    built = {}
    for name, builder in ALL_FIXTURE_BUILDERS.items():
        built[name] = builder()
    return built


@pytest.fixture(params=FIXTURE_NAMES)
def fixture(request, built_fixtures):
    """Parameterized access to each dataset in turn."""
    return built_fixtures[request.param]


def pytest_configure(config):
    config.addinivalue_line("markers", "render: builds a real PDF (slow, needs LibreOffice)")
