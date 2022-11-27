#!/usr/bin/env python
import sys
import argparse
import time

from ooodev.utils.lo import Lo
from select_listener import SelectionListener

def args_add(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "-t",
        "--auto-terminate",
        help="Optional Auto Terminate",
        action="store_true",
        dest="auto_terminate",
        default=False,
    )


def main_loop() -> None:
    # https://stackoverflow.com/a/8685815/1171746
    # create parser to read terminal input
    parser = argparse.ArgumentParser(description="main")

    # add args to parser
    args_add(parser=parser)

    # read the current command line args
    args = parser.parse_args()
    sl = SelectionListener()

    # delay in seconds
    delay = 1.5

    # check an see if user passed in a auto terminate option
    if args.auto_terminate:
        Lo.delay(delay)
        Lo.close_office()
        return

    # stop run min and max to raise listen events

    # while Writer is open, keep running the script unless specifically ended by user
    while 1:
        if sl.closed is True:  # wait for windowClosed event to be raised
            print("\nExiting by document close.\n")
            break
        time.sleep(0.1)


if __name__ == "__main__":
    print("Press 'ctl+c' to exit script early.")
    try:
        main_loop()
    except KeyboardInterrupt:
        # ctrl+c exitst the script earily
        print("\nExiting by user request.\n", file=sys.stderr)
        sys.exit(0)
