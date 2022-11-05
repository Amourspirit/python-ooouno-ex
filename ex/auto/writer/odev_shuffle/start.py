#!/usr/bin/env python
# coding: utf-8
from __future__ import annotations
import argparse
from typing import cast
import random

from ooodev.dialog.msgbox import MsgBox, MessageBoxType, MessageBoxButtonsEnum, MessageBoxResultsEnum
from ooodev.utils.lo import Lo
from ooodev.office.write import Write
from ooodev.utils.gui import GUI

from com.sun.star.text import XTextDocument


def args_add(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "-f",
        "--file",
        help="File path of input file to convert",
        action="store",
        dest="file_path",
        required=True,
    )


def apply_shuffle(doc: XTextDocument, delay: int, visible: bool) -> None:
    doc_text = doc.getText()
    if visible:
        cursor = Write.get_view_cursor(doc)
    else:
        cursor = Write.get_cursor(doc)

    word_cursor = Write.get_word_cursor(doc)
    word_cursor.gotoStart(False)  # go to start of text

    while True:
        word_cursor.gotoNextWord(True)

        # move the text view cursor, and highlight the current word
        cursor.gotoRange(word_cursor.getStart(), False)
        cursor.gotoRange(word_cursor.getEnd(), True)
        curr_word = word_cursor.getString()

        # get whitespace padding amounts
        c_len = len(curr_word)
        curr_word = curr_word.lstrip()
        l_pad = c_len - len(curr_word)  # left whitespace padding amount
        curr_word = curr_word.rstrip()
        r_pad = c_len - len(curr_word) - l_pad  # right whitespace padding ammount
        if len(curr_word) > 0:
            pad_l = " " * l_pad  # recreate left padding
            pad_r = " " * r_pad  # recreate right padding
            Lo.delay(delay)
            mid_shuff = mid_shuffle(curr_word)
            doc_text.insertString(word_cursor, f"{pad_l}{mid_shuff}{pad_r}", True)

        if word_cursor.gotoNextWord(False) is False:
            break

    word_cursor.gotoStart(False)  # go to start of text
    cursor.gotoStart(False)


def mid_shuffle(s: str) -> str:
    s_len = len(s)
    if s_len <= 3:  # not long enough
        return s

    # extract middle of the string for shuffling
    mid = s[1 : s_len - 1]

    # shuffle a list of characters  made from the middle
    mid_lst = list(mid)
    random.shuffle(mid_lst)

    # rebuild string, adding back the first and last letters
    # punctuation may be first or last char
    return f"{s[:1]}{''.join(mid_lst)}{s[-1:]}"


def main() -> int:
    visible = True
    loop_delay = 150

    # create parser to read terminal input
    parser = argparse.ArgumentParser(description="main")

    # add args to parser
    args_add(parser=parser)

    # read the current command line args
    args = parser.parse_args()

    loader = Lo.load_office(Lo.ConnectPipe())

    fnm = cast(str, args.file_path)

    try:
        doc = Write.open_doc(fnm=fnm, loader=loader)
        GUI.set_visible(is_visible=visible, odoc=doc)
        apply_shuffle(doc, loop_delay, visible)

        Lo.delay(1_000)
        msg_result = MsgBox.msgbox(
            "Do you wish to save document?",
            "Save",
            boxtype=MessageBoxType.QUERYBOX,
            buttons=MessageBoxButtonsEnum.BUTTONS_YES_NO,
        )
        if msg_result == MessageBoxResultsEnum.YES:
            Write.save_doc(text_doc=doc, fnm="shuffled.odt")

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
    except Exception:
        Lo.close_office()
        raise

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
