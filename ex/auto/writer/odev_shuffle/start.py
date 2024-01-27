from __future__ import annotations
import sys
import argparse
from typing import cast
from pathlib import Path
import random

from ooodev.dialog.msgbox import (
    MsgBox,
    MessageBoxType,
    MessageBoxButtonsEnum,
    MessageBoxResultsEnum,
)
from ooodev.utils.lo import Lo
from ooodev.write import WriteDoc


def args_add(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "-f",
        "--file",
        help="File path of input file to convert",
        action="store",
        dest="file_path",
        required=True,
    )


def apply_shuffle(doc: WriteDoc, delay: int, visible: bool) -> None:
    doc_text = doc.get_text()
    if visible:
        cursor = doc.get_view_cursor()
    else:
        cursor = doc.get_cursor()

    word_cursor = doc.get_cursor()
    word_cursor.goto_start()  # go to start of text

    while True:
        word_cursor.goto_next_word(True)

        # move the text view cursor, and highlight the current word
        cursor.goto_range(word_cursor.component.getStart())
        cursor.goto_range(word_cursor.component.getEnd(), True)
        curr_word = word_cursor.get_string()

        # get whitespace padding amounts
        c_len = len(curr_word)
        curr_word = curr_word.lstrip()
        l_pad = c_len - len(curr_word)  # left whitespace padding amount
        curr_word = curr_word.rstrip()
        r_pad = c_len - len(curr_word) - l_pad  # right whitespace padding amount
        if len(curr_word) > 0:
            pad_l = " " * l_pad  # recreate left padding
            pad_r = " " * r_pad  # recreate right padding
            Lo.delay(delay)
            mid_shuffle = do_mid_shuffle(curr_word)
            doc_text.insert_string(
                word_cursor.component, f"{pad_l}{mid_shuffle}{pad_r}", True
            )

        if word_cursor.goto_next_word() is False:
            break

    word_cursor.goto_start()  # go to start of text
    cursor.goto_start()


def do_mid_shuffle(s: str) -> str:
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

    if len(sys.argv) == 1:
        pth = Path(__file__).parent / "data" / "cicero_dummy.odt"
        sys.argv.append("-f")
        sys.argv.append(str(pth))

    # read the current command line args
    args = parser.parse_args()

    loader = Lo.load_office(Lo.ConnectPipe())

    fnm = cast(str, args.file_path)

    try:
        doc = WriteDoc.open_doc(fnm=fnm, loader=loader)
        doc.set_visible(visible)
        if visible:
            Lo.delay(300)  # delay for document to load
            doc.zoom()
        apply_shuffle(doc, loop_delay, visible)

        Lo.delay(1_000)
        msg_result = MsgBox.msgbox(
            "Do you wish to save document?",
            "Save",
            boxtype=MessageBoxType.QUERYBOX,
            buttons=MessageBoxButtonsEnum.BUTTONS_YES_NO,
        )
        if msg_result == MessageBoxResultsEnum.YES:
            pth = Path.cwd() / "tmp"
            pth.mkdir(exist_ok=True)
            doc.save_doc(fnm=pth / "shuffled.odt")

        msg_result = MsgBox.msgbox(
            "Do you wish to close document?",
            "All done",
            boxtype=MessageBoxType.QUERYBOX,
            buttons=MessageBoxButtonsEnum.BUTTONS_YES_NO,
        )
        if msg_result == MessageBoxResultsEnum.YES:
            doc.close_doc()
            Lo.close_office()
        else:
            print("Keeping document open")
    except Exception:
        Lo.close_office()
        raise

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
