#!/usr/bin/env python
# coding: utf-8
#
# on wayland (some versions of Linux)
# may get error:
#    (soffice:67106): Gdk-WARNING **: 02:35:12.168: XSetErrorHandler() called with a GDK error trap pushed. Don't do that.
# This seems to be a Wayland/Java compatability issues.
# see: http://www.babelsoft.net/forum/viewtopic.php?t=24545

import sys
import argparse
from typing import Any, cast

from show_sheet import ShowSheet

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.gbl_named_event import GblNamedEvent
from ooodev.events.lo_events import LoEvents


from com.sun.star.util import XProtectable


def args_add(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "-f",
        "--file",
        help="File path of input file to convert",
        action="store",
        dest="file_path",
        required=True,
    )
    parser.add_argument(
        "-o",
        "--out",
        help="Optional file path of output file",
        action="store",
        dest="out_file",
        default="",
    )
    parser.add_argument(
        "-r", "--read-only", help="Read only mode", action="store_true", dest="read_only", default=False
    )
    parser.add_argument("-s", "--show", help="Show Document", action="store_true", dest="show", default=False)
    parser.add_argument("-v", "--verbose", help="Verbose output", action="store_true", dest="verbose", default=False)


def on_lo_print(source: Any, e: CancelEventArgs) -> None:
    # this method is a callback for ooodev internal printing
    # by setting e.canecl = True all internal printing of ooodev is suppressed
    e.cancel = True


def main() -> int:
    # create parser to read terminal input
    parser = argparse.ArgumentParser(description="main")

    # add args to parser
    args_add(parser=parser)

    if len(sys.argv) <= 1:
        parser.print_help()
        return 0

    # read the current command line args
    args = parser.parse_args()

    visible = args.show
    if visible:
        delay = 2_000
    else:
        delay = 0

    if not args.verbose:
        # hook ooodev internal printing event
        LoEvents().on(GblNamedEvent.PRINTING, on_lo_print)

    ss = ShowSheet(read_only=args.read_only, input_fnm=args.file_path, out_fnm=args.out_file, visible=args.show)
    ss.main()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
