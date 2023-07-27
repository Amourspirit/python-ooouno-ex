# Print all the text in every paragraph using enumeration access.

from __future__ import annotations
import sys
import argparse
from typing import cast
from pathlib import Path

import uno
from com.sun.star.text import XText
from com.sun.star.text import XTextContent
from com.sun.star.text import XTextRange

from ooodev.office.write import Write
from ooodev.utils.gui import GUI
from ooodev.utils.info import Info
from ooodev.utils.lo import Lo
from ooodev.utils.props import Props
from ooodev.wrapper.break_context import BreakContext


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


def print_paras(xtext: XText) -> None:
    # iterate through the document contents, printing all the text portions in each paragraph
    try:
        text_enum = Write.get_enumeration(xtext)
        while text_enum.hasMoreElements():
            # return a paragraph (or text table)
            tc = Lo.qi(XTextContent, text_enum.nextElement(), True)

            if not Info.support_service(tc, "com.sun.star.text.TextTable"):
                print("P--")
                para_enum = Write.get_enumeration(tc)
                while para_enum.hasMoreElements():
                    txt_range = Lo.qi(XTextRange, para_enum.nextElement(), True)

                    # return a text portion
                    print(f'  {Props.get_property(txt_range, "TextPortionType")} = "{txt_range.getString()}"')
            else:
                print("Text table")
    except Exception as e:
        print(e)


def main() -> int:
    # create parser to read terminal input
    parser = argparse.ArgumentParser(description="main")

    # add args to parser
    args_add(parser=parser)

    if len(sys.argv) <= 1:
        # parser.print_help()
        # return 0
        pth = Path(__file__).parent / "data" / "cicero_dummy.odt"
        sys.argv.append("--show")
        sys.argv.append("-f")
        sys.argv.append(str(pth))

    # read the current command line args
    args = parser.parse_args()

    visible = args.show

    # Using Lo.Loader context manager wrapped by BreakContext load Office and connect via socket.
    # Context manager takes care of terminating instance when job is done.
    # see: https://python-ooo-dev-tools.readthedocs.io/en/latest/src/wrapper/break_context.html
    # see: https://python-ooo-dev-tools.readthedocs.io/en/latest/src/utils/lo.html#ooodev.utils.lo.Lo.Loader
    with BreakContext(Lo.Loader(connector=Lo.ConnectSocket(), opt=Lo.Options(verbose=args.verbose))) as loader:

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
                GUI.set_visible(visible=visible, doc=doc)
            print_paras(doc.getText())

        finally:
            Lo.close_doc(doc)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
