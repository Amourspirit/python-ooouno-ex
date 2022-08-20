#!/usr/bin/env python
# coding: utf-8
from curses.ascii import isdigit
import sys
import argparse
from typing import Any, cast

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.gbl_named_event import GblNamedEvent
from ooodev.events.lo_events import LoEvents
from ooodev.office.write import Write
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.utils.color import CommonColor, Color
from ooodev.utils.props import Props
from ooodev.wrapper.break_context import BreakContext

from com.sun.star.text import XTextDocument
from com.sun.star.text import XTextRange
from com.sun.star.util import XSearchable

from ooo.dyn.awt.font_slant import FontSlant  # enum


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
    parser.add_argument(
        "--word",
        action="append",
        nargs=2,
        required=True,
        help="Word color pairs where word is the word to italicize and color is a named color such as red or a color integer such as 16711680",
    )


def italicize_all(doc: XTextDocument, phrase: str, color: Color) -> int:
    # cursor = Write.get_view_cursor(doc) # can be used when visible
    cursor = Write.get_cursor(doc)
    cursor.gotoStart(False)
    page_cursor = Write.get_page_cursor(doc)
    result = 0
    try:
        xsearchable = Lo.qi(XSearchable, doc, True)
        srch_desc = xsearchable.createSearchDescriptor()
        print(f"Searching for all occurrences of '{phrase}'")
        pharse_len = len(phrase)
        srch_desc.setSearchString(phrase)
        # for props see: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1util_1_1SearchDescriptor.html
        Props.set_property(obj=srch_desc, name="SearchCaseSensitive", value=False)
        Props.set_property(
            obj=srch_desc, name="SearchWords", value=True
        )  # If TRUE, only complete words will be found.

        matches = xsearchable.findAll(srch_desc)
        result = matches.getCount()

        print(f"No. of matches: {result}")

        for i in range(result):
            match_tr = Lo.qi(XTextRange, matches.getByIndex(i))
            if match_tr is not None:
                cursor.gotoRange(match_tr, False)
                print(f"  - found: '{match_tr.getString()}'")
                print(f"    - on page {page_cursor.getPage()}")
                cursor.gotoStart(True)
                print(f"    - starting at char position: {len(cursor.getString()) - pharse_len}")

                Props.set_properties(obj=match_tr, names=("CharColor", "CharPosture"), vals=(color, FontSlant.ITALIC))

    except Exception as e:
        raise
    return result


def get_color(color: str) -> int:
    if color.isdigit():
        return int(color)
    c = color.upper()
    if hasattr(CommonColor, c):
        return getattr(CommonColor, c)
    return CommonColor.RED


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

            with Lo.ControllerLock():
                for word, color in args.word:
                    result = italicize_all(doc, word, get_color(color))
                    print(f"Found {result} results for {word}")

            Lo.delay(delay)
            Write.save_doc(text_doc=doc, fnm="italicized.doc")

        finally:
            Lo.close_doc(doc)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
