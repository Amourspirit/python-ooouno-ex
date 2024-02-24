import sys
import argparse
from typing import Sequence, cast
from pathlib import Path

import uno

from ooodev.dialog.msgbox import (
    MessageBoxType,
    MessageBoxButtonsEnum,
    MessageBoxResultsEnum,
)
from ooodev.loader import Lo
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
    parser.add_argument(
        "-v",
        "--verbose",
        help="Verbose output",
        action="store_true",
        dest="verbose",
        default=False,
    )


def find_words(doc: WriteDoc, words: Sequence[str]) -> None:
    # get the view cursor and link the page cursor to it
    tvc = doc.get_view_cursor()
    tvc.goto_start()
    search = doc.create_search_descriptor()
    search.search_regular_expression = True
    for word in words:
        search.search_str = word
        result = search.find_first()
        if result is not None:
            tvc.goto_range(result.component)
            print(f"  - found '{result.get_string()}'")
            print(f"    - on page {tvc.get_page()}")
            tvc.go_right(len(result.get_string()), True)
            print(f"    - at char position: {len(tvc.get_string())}")
            Lo.delay(500)


def replace_words(
    doc: WriteDoc, old_words: Sequence[str], new_words: Sequence[str]
) -> int:
    replace = doc.create_replace_descriptor()
    result = replace.replace_words(old_words, new_words)
    print(f"Replaced {result} words")


def main() -> int:
    # create parser to read terminal input
    parser = argparse.ArgumentParser(description="main")

    # add args to parser
    args_add(parser=parser)

    if len(sys.argv) <= 1:
        # parser.print_help()
        # return 0
        pth = Path(__file__).parent / "data" / "bigStory.doc"
        sys.argv.append("-f")
        sys.argv.append(str(pth))

    # read the current command line args
    args = parser.parse_args()

    delay = 4_000

    loader = Lo.load_office(
        connector=Lo.ConnectSocket(), opt=Lo.Options(verbose=args.verbose)
    )

    fnm = cast(str, args.file_path)

    try:
        doc = WriteDoc.open_doc(fnm=fnm, loader=loader, visible=True)
        uk_words = ("colour", "neighbour", "centre", "behaviour", "metre", "through")
        us_words = ("color", "neighbor", "center", "behavior", "meter", "thru")

        words = (
            "(G|g)rit",
            "colou?r",
        )
        find_words(doc, words)

        _ = replace_words(doc, uk_words, us_words)

        Lo.delay(delay)
        msg_result = doc.msgbox(
            "Do you wish to save document?",
            "Save",
            boxtype=MessageBoxType.QUERYBOX,
            buttons=MessageBoxButtonsEnum.BUTTONS_YES_NO,
        )
        if msg_result == MessageBoxResultsEnum.YES:
            pth = Path.cwd() / "tmp"
            pth.mkdir(parents=True, exist_ok=True)
            doc.save_doc(fnm=pth / "replaced.doc")

        msg_result = doc.msgbox(
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
