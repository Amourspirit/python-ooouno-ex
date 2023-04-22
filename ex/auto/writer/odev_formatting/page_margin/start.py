from __future__ import annotations
from typing import cast
import sys
import argparse

import uno
from ooodev.dialog.msgbox import MsgBox, MessageBoxType, MessageBoxButtonsEnum, MessageBoxResultsEnum
from ooodev.format.writer.modify.page.page import Margins
from ooodev.office.write import Write
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo


def args_add(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "-f",
        "--file",
        help="File path of input file to style",
        action="store",
        dest="file_path",
        required=True,
    )
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

    # resources\odt\cicero_dummy.odt
    fnm = cast(str, args.file_path)

    # start LibreOffice
    _ = Lo.load_office(Lo.ConnectPipe(), opt=Lo.Options(verbose=args.verbose))

    try:
        doc = Write.open_doc(fnm=fnm)

        # show the document
        GUI.set_visible()
        GUI.zoom(GUI.ZoomEnum.OPTIMAL)
        Lo.delay(4_000) # give user time to see the document before styling.

        # create a style margin object
        style = Margins(left=30, right=30, top=40, bottom=18, gutter=8)
        # apply margin style to the document
        style.apply(doc)

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
    SystemExit(main())
