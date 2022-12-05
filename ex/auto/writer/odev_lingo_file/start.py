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
from typing import cast

import uno
from com.sun.star.text import XTextDocument

from ooodev.office.write import Write
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.wrapper.break_context import BreakContext


def check_sentences(doc: XTextDocument) -> None:
    # load spell checker, proof reader
    speller = Write.load_spell_checker()
    proofreader = Write.load_proofreader()

    para_cursor = Write.get_paragraph_cursor(doc)
    para_cursor.gotoStart(False)  # go to start test; no selection

    while 1:
        para_cursor.gotoEndOfParagraph(True)  # select all of paragraph
        curr_para_str = para_cursor.getString()

        if len(curr_para_str) > 0:
            print(f"\n>> {curr_para_str}")

            sentences = Write.split_paragraph_into_sentences(curr_para_str)
            for sentence in sentences:
                # print(f'S <{sentence}>')
                Write.proof_sentence(sentence, proofreader)
                Write.spell_sentence(sentence, speller)

        if para_cursor.gotoNextParagraph(False) is False:
            break


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

    # Using Lo.Loader context manager wraped by BreakContext load Office and connect via socket.
    # Context manager takes care of terminating instance when job is done.
    # see: https://python-ooo-dev-tools.readthedocs.io/en/latest/src/wrapper/break_context.html
    # see: https://python-ooo-dev-tools.readthedocs.io/en/latest/src/utils/lo.html#ooodev.utils.lo.Lo.Loader
    with BreakContext(
        Lo.Loader(connector=Lo.ConnectSocket(headless=not visible), opt=Lo.Options(verbose=args.verbose))
    ) as loader:

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

            check_sentences(doc)

            Lo.delay(delay)

        finally:
            Lo.close_doc(doc)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
