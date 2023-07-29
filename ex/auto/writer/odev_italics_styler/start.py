from __future__ import annotations
import sys
import argparse
from typing import cast
from pathlib import Path

import uno
from com.sun.star.text import XTextDocument
from com.sun.star.text import XTextRange
from com.sun.star.util import XSearchable

from ooodev.dialog.msgbox import MsgBox, MessageBoxType, MessageBoxButtonsEnum, MessageBoxResultsEnum
from ooodev.office.write import Write
from ooodev.utils.color import CommonColor, Color
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.utils.props import Props
from ooodev.format.writer.direct.char.font import Font


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
        searchable = Lo.qi(XSearchable, doc, True)
        search_desc = searchable.createSearchDescriptor()
        print(f"Searching for all occurrences of '{phrase}'")
        phrase_len = len(phrase)
        search_desc.setSearchString(phrase)
        # for props see: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1util_1_1SearchDescriptor.html
        Props.set_property(obj=search_desc, name="SearchCaseSensitive", value=False)
        Props.set_property(
            obj=search_desc, name="SearchWords", value=True
        )  # If TRUE, only complete words will be found.

        matches = searchable.findAll(search_desc)
        result = matches.getCount()

        print(f"No. of matches: {result}")

        font_effect = Font(i=True, color=color)

        for i in range(result):
            match_tr = Lo.qi(XTextRange, matches.getByIndex(i))
            if match_tr is not None:
                cursor.gotoRange(match_tr, False)
                print(f"  - found: '{match_tr.getString()}'")
                print(f"    - on page {page_cursor.getPage()}")
                cursor.gotoStart(True)
                print(f"    - starting at char position: {len(cursor.getString()) - phrase_len}")

                font_effect.apply(match_tr)

    except Exception:
        raise
    return result


def get_color(color: str) -> int:
    try:
        return CommonColor.from_str(color)
    except Exception:
        pass
    return CommonColor.RED


def main() -> int:
    # create parser to read terminal input
    parser = argparse.ArgumentParser(description="main")

    # add args to parser
    args_add(parser=parser)

    if len(sys.argv) <= 1:
        # parser.print_help()
        # return 0
        pth = Path(__file__).parent / "data" / "cicero_dummy.odt"
        sys.argv.append("-f")
        sys.argv.append(str(pth))
        sys.argv.append("--word")
        sys.argv.append("pleasure")
        sys.argv.append("green")
        sys.argv.append("--word")
        sys.argv.append("pain")
        sys.argv.append("red")

    # read the current command line args
    args = parser.parse_args()

    delay = 3_000

    loader = Lo.load_office(connector=Lo.ConnectSocket(), opt=Lo.Options(verbose=args.verbose))

    fnm = cast(str, args.file_path)

    try:
        doc = Write.open_doc(fnm=fnm, loader=loader)

        GUI.set_visible(visible=True, doc=doc)

        with Lo.ControllerLock():
            for word, color in args.word:
                result = italicize_all(doc, word, get_color(color))
                print(f"Found {result} results for {word}")

        Lo.delay(delay)
        msg_result = MsgBox.msgbox(
            "Do you wish to save document?",
            "Save",
            boxtype=MessageBoxType.QUERYBOX,
            buttons=MessageBoxButtonsEnum.BUTTONS_YES_NO,
        )
        if msg_result == MessageBoxResultsEnum.YES:
            pth = Path.cwd() / "tmp"
            pth.mkdir(exist_ok=True)
            Write.save_doc(text_doc=doc, fnm=pth / "italicized.doc")

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
