#!/usr/bin/env python
# coding: utf-8
# region Imports
from __future__ import annotations
import argparse
import sys
from typing import Any

from ooodev.events.args.dispatch_args import DispatchArgs
from ooodev.events.args.dispatch_cancel_args import DispatchCancelArgs
from ooodev.events.lo_events import Events
from ooodev.events.lo_named_event import LoNamedEvent

from dispatcher import Dispatcher

# endregion Imports

# region args
def args_add(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "-d",
        "--doc",
        help="Path to document",
        action="store",
        dest="fnm_doc",
        required=True,
    )


# endregion args

# region dispatch events
def on_dispatching(source: Any, event: DispatchCancelArgs) -> None:
    if event.cmd == "About":
        event.cancel = True
        return


def on_dispatched(source: Any, event: DispatchArgs) -> None:
    print(f"Dispatched: {event.cmd}")


# endregion dispatch events

# region Main
def main() -> int:
    # create parser to read terminal input
    parser = argparse.ArgumentParser(description="main")

    # add args to parser
    args_add(parser=parser)

    if len(sys.argv) <= 1:
        parser.print_help(sys.stderr)
        return 1

    # read the current command line args
    args = parser.parse_args()

    fnm = args.fnm_doc
    events = Events()
    events.on(LoNamedEvent.DISPATCHING, on_dispatching)
    events.on(LoNamedEvent.DISPATCHED, on_dispatched)

    dpatch = Dispatcher(fnm)
    dpatch.main()
    return 0

if __name__ == "__main__":
    raise SystemExit(main())

# endregion main
