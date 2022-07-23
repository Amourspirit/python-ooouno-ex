#!/usr/bin/env python
# coding: utf-8
# region Imports
from __future__ import annotations

from ooodev.utils.lo import Lo
from ooodev.exceptions import ex as LoEx
from ooodev.utils.session import Session

# endregion Imports

def main() -> int:
    try:
        import pyglobal
    except ImportError:
        print("As expected unable to import pyglobal without registering path.")
        print()
    try:
        Session.register_path(Session.PathEnum.SHARE_USER_PYTHON)
    except LoEx.LoNotLoadedError:
        print("As expected unable to register path before Lo.load_office is called")
        print()

    with Lo.Loader(Lo.ConnectPipe()) as loader:
        Session.register_path(Session.PathEnum.SHARE_USER_PYTHON)
    
    # Even though the connection to office is now closed the path is registered in sys.path
    # and the module can still be imported.
    
    # pyglobal is a module in 'My Macros'
    # pyglobal cannot be imported by default as the 'My Macros' directory is not known to python.
    # outside of a macro office must be loaded before 'My Macros' directory can be resolved.
    # Session.register_path(Session.PathEnum.SHARE_USER_PYTHON) takes care of resolving the path and 
    # adding it to python's sys.path.
    import pyglobal
    print(pyglobal.G_COUNT)
    for _ in range(10):
        pyglobal.G_COUNT += 1
        print(pyglobal.G_COUNT)

    return 0

if __name__ == "__main__":
    raise SystemExit(main())

