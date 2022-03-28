# coding: utf-8
import pytest
from pathlib import Path

@pytest.fixture(scope="session")
def fixture_path(app_root: Path) -> Path:
    return app_root / 'tests' / 'fixture'

