# coding: utf-8
import os
import sys
import __main__
from pathlib import Path
from typing import List
import config

_APP_ROOT = None
_OS_PATH_SET = False
_APP_CFG = None


def get_root() -> str:
    """
    Gets Application Root Path

    Returns:
        str: App root as string.
    """
    global _APP_ROOT
    if _APP_ROOT is None:
        _APP_ROOT = os.environ.get(
            'project_root', str(Path(__main__.__file__).parent))
    return _APP_ROOT

def set_os_root_path() -> None:
    """
    Ensures application root dir is in sys.path
    """
    global _OS_PATH_SET
    if _OS_PATH_SET is False:
        _app_root = get_root()
        if not _app_root in sys.path:
            sys.path.insert(0, _app_root)
    _OS_PATH_SET = True

def get_app_cfg() -> config.AppConfig:
    """
    Get App Config. config is cached
    """
    global _APP_CFG
    if _APP_CFG is None:
        _APP_CFG = config.read_config_default()
    return _APP_CFG

def get_path_from_lst(lst: List[str], ensure_absolute: bool = False) -> Path:
    """
    Builds a Path from a list of strings
    
    If lsg[0] starts with ``~`` then it is expanded to user home dir.

    Args:
        lst (List[str]): List of path parts
        ensure_absolute (bool, optional): If true returned will have root dir prepended
            if path is not absolute

    Raises:
        ValueError: If lst is empty

    Returns:
        Path: Path of combined from ``lst``
    """
    if len(lst) == 0:
        raise ValueError("lst arg is zero length")
    arg = lst[0]
    expand = arg.startswith('~')
    p = Path(*lst)
    if expand:
        p = p.expanduser()
    if ensure_absolute is True and p.is_absolute() is False:
        p = Path(get_root(), p)
    return p