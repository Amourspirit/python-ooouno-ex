#!/usr/bin/env python
# coding: utf-8
from __future__ import annotations
import argparse
from typing import cast
import os

from ooodev.utils.lo import Lo
from ooodev.office.write import Write
from ooodev.utils.info import Info
from ooodev.utils.gui import GUI
from ooodev.wrapper.break_context import BreakContext

from com.sun.star.text import XTextDocument
from com.sun.star.view import XLineCursor


def args_add(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "-f",
        "--file",
        help="File path of input file to convert",
        action="store",
        dest="file_path",
        required=True,
    )


def show_paragraphs(doc: XTextDocument) -> None:
    tvc = Write.get_view_cursor(doc)
    para_cursor = Write.get_paragraph_cursor(doc)
    para_cursor.gotoStart(False)  # go to start test; no selection

    while 1:
        para_cursor.gotoEndOfParagraph(True)  # select all of paragraph
        curr_para = para_cursor.getString()
        if len(curr_para) > 0:
            tvc.gotoRange(para_cursor.getStart(), False)
            tvc.gotoRange(para_cursor.getEnd(), True)

            print(f"P<{curr_para}>")
            Lo.delay(500)  # delay half a second

        if para_cursor.gotoNextParagraph(False) is False:
            break


def count_words(doc: XTextDocument) -> int:
    word_cursor = Write.get_word_cursor(doc)
    word_cursor.gotoStart(False)  # go to start of text

    word_count = 0
    while 1:
        word_cursor.gotoEndOfWord(True)
        curr_word = word_cursor.getString()
        if len(curr_word) > 0:
            word_count += 1
        if word_cursor.gotoNextWord(False) is False:
            break
    return word_count


def show_lines(doc: XTextDocument) -> None:
    tvc = Write.get_view_cursor(doc)
    tvc.gotoStart(False)  # go to start of text

    line_cursor = Lo.qi(XLineCursor, tvc, True)
    have_text = True
    while have_text is True:
        line_cursor.gotoStartOfLine(False)
        line_cursor.gotoEndOfLine(True)
        print(f"L<{tvc.getString()}>")  # no text selection in line cursor
        Lo.delay(500)  # delay half a second
        tvc.collapseToEnd()
        have_text = tvc.goRight(1, True)


def main() -> int:
    # create parser to read terminal input
    parser = argparse.ArgumentParser(description="main")

    # add args to parser
    args_add(parser=parser)

    # read the current command line args
    args = parser.parse_args()

    # Using Lo.Loader context manager wraped by BreakContext load Office and connect via socket.
    # Context manager takes care of terminating instance when job is done.
    # Note the use of the headless flag. Not using GUI for process.
    # see: https://python-ooo-dev-tools.readthedocs.io/en/latest/src/wrapper/break_context.html
    # see: https://python-ooo-dev-tools.readthedocs.io/en/latest/src/utils/lo.html#ooodev.utils.lo.Lo.Loader
    with BreakContext(Lo.Loader(Lo.ConnectSocket())) as loader:

        fnm = cast(str, args.file_path)

        try:
            doc = Write.open_doc(fnm=fnm, loader=loader)
        except Exception:
            print(f"Could not open '{fnm}'")
            # office will close and with statement is exited
            raise BreakContext.Break

        try:
            GUI.set_visible(is_visible=True, odoc=doc)

            show_paragraphs(doc)
            print(f"Word count: {count_words(doc)}")
            show_lines(doc)
        finally:
            Lo.close_doc(doc)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
