import sys
import argparse
from typing import Sequence, cast
from pathlib import Path

import uno
from com.sun.star.beans import XPropertySet
from com.sun.star.text import XTextDocument
from com.sun.star.text import XTextRange
from com.sun.star.util import XReplaceable
from com.sun.star.util import XReplaceDescriptor
from com.sun.star.util import XSearchable

from ooodev.dialog.msgbox import MsgBox, MessageBoxType, MessageBoxButtonsEnum, MessageBoxResultsEnum
from ooodev.office.write import Write
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo


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
            print(f"    - at char position: {len(tvc.getString())}")
            Lo.delay(500)


def replace_words(doc: XTextDocument, old_words: Sequence[str], new_words: Sequence[str]) -> int:
    replaceable = Lo.qi(XReplaceable, doc, raise_err=True)
    replace_desc = Lo.qi(XReplaceDescriptor, replaceable.createSearchDescriptor())

    for old, new in zip(old_words, new_words):
        replace_desc.setSearchString(old)
        replace_desc.setReplaceString(new)
    return replaceable.replaceAll(replace_desc)


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

    loader = Lo.load_office(connector=Lo.ConnectSocket(), opt=Lo.Options(verbose=args.verbose))

    fnm = cast(str, args.file_path)

    try:
        doc = Write.open_doc(fnm=fnm, loader=loader)
        uk_words = ("colour", "neighbour", "centre", "behaviour", "metre", "through")
        us_words = ("color", "neighbor", "center", "behavior", "meter", "thru")

        GUI.set_visible(visible=True, doc=doc)

        words = (
            "(G|g)rit",
            "colou?r",
        )
        find_words(doc, words)

        _ = replace_words(doc, uk_words, us_words)

        Lo.delay(delay)
        msg_result = MsgBox.msgbox(
            "Do you wish to save document?",
            "Save",
            boxtype=MessageBoxType.QUERYBOX,
            buttons=MessageBoxButtonsEnum.BUTTONS_YES_NO,
        )
        if msg_result == MessageBoxResultsEnum.YES:
            pth = Path.cwd() / "tmp"
            pth.mkdir(parents=True, exist_ok=True)
            Write.save_doc(text_doc=doc, fnm=pth / "replaced.doc")

        msg_result = MsgBox.msgbox(
            "Do you wish to close document?",
            "All done",
            boxtype=MessageBoxType.QUERYBOX,
            buttons=MessageBoxButtonsEnum.BUTTONS_YES_NO,
        )
        if msg_result == MessageBoxResultsEnum.YES:
            Lo.close_doc(doc=doc)
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
