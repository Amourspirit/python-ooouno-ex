#!/usr/bin/env python
# coding: utf-8
import sys
import argparse
from typing import Any, cast
from pathlib import Path

import uno

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.gbl_named_event import GblNamedEvent
from ooodev.events.lo_events import LoEvents
from ooodev.office.write import Write
from ooodev.utils.gui import GUI
from ooodev.utils.images_lo import ImagesLo
from ooodev.utils.lo import Lo
from ooodev.utils.props import Props
from ooodev.wrapper.break_context import BreakContext

from com.sun.star.awt import Size
from com.sun.star.text import XTextShapesSupplier


def args_add(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "-f",
        "--file",
        help="File path of input file to convert",
        action="store",
        dest="file_path",
        required=True,
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
        delay = 4_000
    else:
        delay = 0

    if not args.verbose:
        # hook ooodev internal printing event
        LoEvents().on(GblNamedEvent.PRINTING, on_lo_print)

    # Using Lo.Loader context manager wraped by BreakContext load Office and connect via socket.
    # Context manager takes care of terminating instance when job is done.
    # see: https://python-ooo-dev-tools.readthedocs.io/en/latest/src/wrapper/break_context.html
    # see: https://python-ooo-dev-tools.readthedocs.io/en/latest/src/utils/lo.html#ooodev.utils.lo.Lo.Loader
    with BreakContext(Lo.Loader(Lo.ConnectSocket(headless=not visible))) as loader:

        fnm = cast(str, args.file_path)

        try:
            doc = Write.open_doc(fnm=fnm, loader=loader)
        except Exception as e:
            print(f"Could not open '{fnm}'")
            print(f"  {e}")
            # office will close and with statement is exited
            raise BreakContext.Break

        try:
            if visible:
                GUI.set_visible(is_visible=visible, odoc=doc)
            Lo.delay(delay)

            pics = Write.get_text_graphics(doc)
            print(f"Num. of text graphics: {len(pics)}")

            # save text graphics to files
            for i, pic in enumerate(pics):
                fnm = Path(f"graphic{i + 1}.png")
                ImagesLo.save_graphic(pic, fnm, "png")  # ".gif", "gif")
                sz = cast(Size, Props.get_property(pic, "SizePixel"))
                print(f"Image size in pixels: {sz.Width} x {sz.Height}")
            print()

            # this supplier is not created; qi() returns None
            shps_supplier = Lo.qi(XTextShapesSupplier, doc)
            if shps_supplier is None:
                print("Could not obtain text shapes supplier")
            else:
                print(f"Number of shapes: {shps_supplier.getShapes().getCount()}")

            # shapes will be implemented at a later date.

            # draw_page = Write.get_shapes(doc)
            # print(draw_page)
            # index_acc = Lo.qi(XIndexAccess, draw_page)
            # print(index_acc.getCount())
            # for i in range(index_acc.getCount()):
            #     print(i, ": ", index_acc.getByIndex(i))
        finally:
            Lo.close_doc(doc)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
