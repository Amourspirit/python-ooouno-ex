#!/usr/bin/env python
# region Imports
from __future__ import annotations
import time
import sys

from doc_monitor import DocMonitor
from ooodev.utils.lo import Lo

# endregion Imports

# region main


def main_loop() -> None:
    # https://stackoverflow.com/a/8685815/1171746
    dw = DocMonitor()

    # check an see if user passed in a auto terminate option
    if len(sys.argv) > 1:
        if str(sys.argv[1]).casefold() in ("t", "true", "y", "yes"):
            Lo.delay(3000)
            Lo.close_office()
            return

    # while Writer is open, keep running the script unless specifically ended by user
    while 1:
        if dw.closed is True:  # wait for windowClosed event to be raised
            print("\nExiting by document close.\n")
            break
        if dw.bridge_disposed is True:
            print("\nExiting due to office bridge is gone\n")
            raise SystemExit(1)
        time.sleep(0.1)


if __name__ == "__main__":
    print("Press 'ctl+c' to exit script early.")
    try:
        main_loop()
    except SystemExit as e:
        sys.exit(e.code)
    except KeyboardInterrupt:
        # ctrl+c exitst the script earily
        print("\nExiting by user request.\n", file=sys.stderr)
        sys.exit(0)

# endregion main
