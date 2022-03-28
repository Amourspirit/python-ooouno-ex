# coding: utf-8
import os
import pytest
from pathlib import Path

def set_app_root():
    p = Path(__file__).parent
    os.environ['project_root'] = str(p)

# set path at this top level. This allows import that call util.get_root() to get proper path.
set_app_root()

@pytest.fixture(scope='session')
def app_root() -> Path:
    return Path(__file__).parent
