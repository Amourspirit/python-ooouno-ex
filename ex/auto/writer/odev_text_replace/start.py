#!/usr/bin/env python
# coding: utf-8
import sys
import argparse
from typing import Any, Sequence, cast

import uno

from ooodev.dialog.msgbox import MsgBox, MessageBoxType, MessageBoxButtonsEnum, MessageBoxResultsEnum
from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.gbl_named_event import GblNamedEvent
from ooodev.events.lo_events import LoEvents
from ooodev.office.write import Write
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo

from com.sun.star.beans import XPropertySet
from com.sun.star.text import XTextDocument
from com.sun.star.text import XTextRange
from com.sun.star.util import XReplaceable
from com.sun.star.util import XReplaceDescriptor
from com.sun.star.util import XSearchable


def args_add(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "-f",
        "--file",
        help="File path of input file to convert",
        action="store",
        dest="file_path",
        required=True,
    )
    parser.add_argument("-v", "--verbose", help="Verbose output", action="store_true", dest="verbose", default=False)


def find_words(doc: XTextDocument, words: Sequence[str]) -> None:
    # get the view cursor and link the page cursor to it
    tvc = Write.get_view_cursor(doc)
    tvc.gotoStart(False)
    page_cursor = Write.get_page_cursor(tvc)
    searchable = Lo.qi(XSearchable, doc)
    srch_desc = searchable.createSearchDescriptor()

    for word in words:
        print(f"Searching for fist occurrence of '{word}'")
        srch_desc.setSearchString(word)

        srch_props = Lo.qi(XPropertySet, srch_desc, raise_err=True)
        srch_props.setPropertyValue("SearchRegularExpression", True)

        srch = searchable.findFirst(srch_desc)

        if srch is not None:
            match_tr = Lo.qi(XTextRange, srch)

            tvc.gotoRange(match_tr, False)
            print(f"  - found '{match_tr.getString()}'")
            print(f"    - on page {page_cursor.getPage()}")
            # tvc.gotoStart(True)
            tvc.goRight(len(match_tr.getString()), True)
            print(f"    - at char postion: {len(tvc.getString())}")
            Lo.delay(500)


def replace_words(doc: XTextDocument, old_words: Sequence[str], new_words: Sequence[str]) -> int:
    replaceable = Lo.qi(XReplaceable, doc, raise_err=True)
    replace_desc = Lo.qi(XReplaceDescriptor, replaceable.createSearchDescriptor())

    for old, new in zip(old_words, new_words):
        replace_desc.setSearchString(old)
        replace_desc.setReplaceString(new)
    return replaceable.replaceAll(replace_desc)


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

    delay = 4_000

    if not args.verbose:
        # hook ooodev internal printing event
        LoEvents().on(GblNamedEvent.PRINTING, on_lo_print)

    loader = Lo.load_office(Lo.ConnectSocket())

    fnm = cast(str, args.file_path)

    try:
        doc = Write.open_doc(fnm=fnm, loader=loader)
        uk_words = ("colour", "neighbour", "centre", "behaviour", "metre", "through")
        us_words = ("color", "neighbor", "center", "behavior", "meter", "thru")

        GUI.set_visible(is_visible=True, odoc=doc)

        words = (
            "(G|g)rit",
            "colou?r",
        )
        find_words(doc, words)

        num = replace_words(doc, uk_words, us_words)

        Lo.delay(delay)
        msg_result = MsgBox.msgbox(
            "Do you wish to save document?",
            "Save",
            boxtype=MessageBoxType.QUERYBOX,
            buttons=MessageBoxButtonsEnum.BUTTONS_YES_NO,
        )
        if msg_result == MessageBoxResultsEnum.YES:
            Write.save_doc(text_doc=doc, fnm="replaced.doc")

        msg_result = MsgBox.msgbox(
            "Do you wish to close document?",
            "All done",
            boxtype=MessageBoxType.QUERYBOX,
            buttons=MessageBoxButtonsEnum.BUTTONS_YES_NO,
        )
        if msg_result == MessageBoxResultsEnum.YES:
            Lo.close_doc(doc=doc, deliver_ownership=True)
            Lo.close_office()
        else:
            print("Keeping document open")
    except Exception as e:
        print(e)
        Lo.close_office()
        raise

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
