from __future__ import annotations
import argparse
import sys
from typing import cast
from pathlib import Path

import uno

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


def show_paragraphs(doc: WriteDoc) -> None:
    tvc = doc.get_view_cursor()
    cursor = doc.get_cursor()
    cursor.goto_start(False)  # go to start test; no selection

    while 1:
        cursor.goto_end_of_paragraph(True)  # select all of paragraph
        curr_para = cursor.get_string()
        if len(curr_para) > 0:
            tvc.goto_range(cursor.component.getStart())
            tvc.goto_range(cursor.component.getEnd(), True)

            print(f"P<{curr_para}>")
            Lo.delay(500)  # delay half a second

        if cursor.goto_next_paragraph() is False:
            break


def count_words(doc: WriteDoc) -> int:
    cursor = doc.get_cursor()
    cursor.goto_start()  # go to start of text

    word_count = 0
    while 1:
        cursor.goto_end_of_word(True)
        curr_word = cursor.get_string()
        if len(curr_word) > 0:
            word_count += 1
        if cursor.goto_next_word() is False:
            break
    return word_count


def show_lines(doc: WriteDoc) -> None:
    tvc = doc.get_view_cursor()
    tvc.goto_start()  # go to start of text

    have_text = True
    while have_text is True:
        tvc.goto_start_of_line()
        tvc.goto_end_of_line(True)
        print(f"L<{tvc.get_string()}>")  # no text selection in line cursor
        Lo.delay(500)  # delay half a second
        tvc.collapse_to_end()
        have_text = tvc.go_right(1, True)


def main() -> int:
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

    loader = Lo.load_office(Lo.ConnectSocket())

    fnm = cast(str, args.file_path)

    try:
        doc = WriteDoc.open_doc(fnm=fnm, loader=loader, visible=True)

        show_paragraphs(doc)
        print(f"Word count: {count_words(doc)}")
        show_lines(doc)

        Lo.delay(1_000)
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
    except Exception as e:
        print(e)
        Lo.close_office()
        raise

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
