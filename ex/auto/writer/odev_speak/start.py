#!/usr/bin/env python
# coding: utf-8
from __future__ import annotations
import argparse
from typing import cast
from pathlib import Path
import re
import sys


def register_proj_path() -> None:
    def get_root_path(pth) -> Path:
        print("testing Path:", pth)
        p = Path(pth)
        for file in p.glob(".root_token"):
            if file.name == ".root_token":
                return file.parent
        parent = p.parent
        if parent == p:
            raise Exception("Got all the way to root. Did not find project root path.")
        return get_root_path(parent)

    ps = str(get_root_path(Path(__file__).absolute()))
    if not ps in sys.path:
        sys.path.insert(0, ps)


# register_proj_path()


from text_to_speech import speak

from ooodev.utils.lo import Lo
from ooodev.office.write import Write
from ooodev.utils.gui import GUI
from ooodev.wrapper.break_context import BreakContext

from com.sun.star.text import XTextDocument
from com.sun.star.text import XSentenceCursor
from com.sun.star.text import XTextRangeCompare


regex = re.compile("[^a-zA-Z0-9, ]")


def args_add(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "-f",
        "--file",
        help="File path of input file to convert",
        action="store",
        dest="file_path",
        required=True,
    )


def speak_sentences(doc: XTextDocument) -> None:
    tvc = Write.get_view_cursor(doc)
    para_cursor = Write.get_paragraph_cursor(doc)
    para_cursor.gotoStart(False)  # go to start test; no selection

    while 1:
        para_cursor.gotoEndOfParagraph(True)  # select all of paragraph
        end_para = para_cursor.getEnd()
        curr_para_str = para_cursor.getString()
        print(f"P<{curr_para_str}>")

        if len(curr_para_str) > 0:
            # set sentence cursor pointing to start of this paragraph
            cursor = para_cursor.getText().createTextCursorByRange(para_cursor.getStart())
            sc = Lo.qi(XSentenceCursor, cursor)
            sc.gotoStartOfSentence(False)
            while 1:
                sc.gotoEndOfSentence(True)  # select all of sentence

                # move the text view cursor to highlight the current sentence
                tvc.gotoRange(sc.getStart(), False)
                tvc.gotoRange(sc.getEnd(), True)
                curr_sent_str = strip_non_word_chars(sc.getString())
                print(f"S<{curr_sent_str}>")
                if len(curr_sent_str) > 0:
                    speak(
                        curr_sent_str,
                    )
                if Write.compare_cursor_ends(sc.getEnd(), end_para) >= Write.CompareEnum.EQUAL:
                    print("Sentence cursor passed end of current paragraph")
                    break

                if sc.gotoNextSentence(False) is False:
                    print("# No next sentence")
                    break

        if para_cursor.gotoNextParagraph(False) is False:
            break


def strip_non_word_chars(s: str) -> str:
    return regex.sub("", s)


def main() -> int:
    # speak("Hello World", save=True, file="/tmp/abc.mp3", speak=True)
    # return
    # create parser to read terminal input
    parser = argparse.ArgumentParser(description="main")

    # add args to parser
    args_add(parser=parser)

    # read the current command line args
    args = parser.parse_args()

    # Using Lo.Loader context manager wraped by BreakContext load Office and connect via socket.
    # Context manager takes care of terminating instance when job is done.
    # see: https://python-ooo-dev-tools.readthedocs.io/en/latest/src/wrapper/break_context.html
    # see: https://python-ooo-dev-tools.readthedocs.io/en/latest/src/utils/lo.html#ooodev.utils.lo.Lo.Loader
    with BreakContext(Lo.Loader(Lo.ConnectSocket())) as loader:

        fnm = cast(str, args.file_path)

        try:
            doc = Write.open_doc(fnm=fnm, loader=loader)
        except Exception as e:
            print(f"Could not open '{fnm}'")
            print(f"  {e}")
            # office will close and with statement is exited
            raise BreakContext.Break

        try:
            GUI.set_visible(is_visible=True, odoc=doc)
            speak_sentences(doc)

        finally:
            Lo.close_doc(doc)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
