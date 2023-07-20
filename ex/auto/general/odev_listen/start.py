#!/usr/bin/env python
# coding: utf-8
# region Imports
from __future__ import annotations
import time
import sys
import argparse

from ooodev.utils.lo import Lo
from ooodev.utils.gui import GUI
from doc_window import DocWindow
from doc_window_adapter import DocWindowAdapter
# endregion Imports


def args_add(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "-t",
        "--auto-terminate",
        help="Optional Auto Terminate",
        action="store_true",
        dest="auto_terminate",
        default=False,
    )
    parser.add_argument(
        "-a",
        "--use-adapter",
        help="Optionally use adapter class",
        action="store_true",
        dest="use_adapter",
        default=False,
    )


# region main


def main_loop() -> None:
    # https://stackoverflow.com/a/8685815/1171746
    
    # create parser to read terminal input
    parser = argparse.ArgumentParser(description="main")

    # add args to parser
    args_add(parser=parser)

    # read the current command line args
    args = parser.parse_args()
    if args.use_adapter:
        dw = DocWindowAdapter()
    else:
        dw = DocWindow()

    # delay in seconds
    delay = 1.5

    # start run min and max to raise listen events
    time.sleep(delay) # wait delay amount of seconds
    for _ in range(3):
        time.sleep(delay)
        GUI.minimize(dw.doc)
        time.sleep(delay)
        GUI.maximize(dw.doc)
    
    # check an see if user passed in a auto terminate option
    if args.auto_terminate:
        Lo.delay(delay)
        Lo.close_office()
        return

    # stop run min and max to raise listen events

    # while Writer is open, keep running the script unless specifically ended by user
    while 1:
        if dw.closed is True: # wait for windowClosed event to be raised
            print("\nExiting by document close.\n")
            break
        time.sleep(0.1)


if __name__ == "__main__":
    print("Press 'ctl+c' to exit script early.")
    try:
        main_loop()
    except KeyboardInterrupt:
        # ctrl+c exists the script early
        print("\nExiting by user request.\n", file=sys.stderr)
        sys.exit(0)

# endregion main
