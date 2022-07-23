#!/usr/bin/env python
# coding: utf-8
# region Imports
from __future__ import annotations
import argparse
import sys
from typing import Any

from ooodev.utils.lo import Lo
from ooodev.utils.gui import GUI
from ooodev.events.lo_events import Events
from ooodev.events.lo_named_event import LoNamedEvent
from ooodev.events.args.dispatch_args import DispatchArgs
from ooodev.events.args.dispatch_cancel_args import DispatchCancelArgs

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
        print("About dispatch canceled")
        event.cancel = True
        return
    print(f"Dispatching: {event.cmd}")


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
    loader = Lo.load_office(Lo.ConnectPipe())
    try:
        doc = Lo.open_doc(fnm=fnm, loader=loader)
        # breakpoint()
    except Exception:
        print(f"Could not open '{fnm}'")
        Lo.close_office()
        raise SystemExit(1)

    # create an instance of events to hook into ooodev events
    events = Events()
    events.on(LoNamedEvent.DISPATCHING, on_dispatching)
    events.on(LoNamedEvent.DISPATCHED, on_dispatched)

    GUI.set_visible(is_visible=True, odoc=doc)
    Lo.delay(3000)  # delay 3 seconds

    # put doc into readonly mode
    Lo.dispatch_cmd("ReadOnlyDoc")
    Lo.delay(1000)

    # opens get involved webpage of LibreOffice in local browser
    Lo.dispatch_cmd("GetInvolved")
    Lo.dispatch_cmd("About")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())

# endregion main
