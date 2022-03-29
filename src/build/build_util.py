# coding: utf-8
from pathlib import Path
from ..utils import util


def get_build_path() -> Path:
    """
    Gets path to build directory

    Returns:
        Path: build directory path
    """
    config = util.get_app_cfg()
    root = Path(util.get_root())
    return root / util.get_path(config.app_build_dir)