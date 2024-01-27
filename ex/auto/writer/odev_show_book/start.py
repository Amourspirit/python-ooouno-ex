# Print all the text in every paragraph using enumeration access.

from __future__ import annotations
import sys
import argparse
from typing import cast
from pathlib import Path

import uno

from ooodev.write import WriteDoc
from ooodev.utils.lo import Lo
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
    parser.add_argument(
        "-s",
        "--show",
        help="Show Document",
        action="store_true",
        dest="show",
        default=False,
    )
    parser.add_argument(
        "-v",
        "--verbose",
        help="Verbose output",
        action="store_true",
        dest="verbose",
        default=False,
    )


def print_paras(doc: WriteDoc) -> None:
    # iterate through the document contents, printing all the text portions in each paragraph
    try:
        text_enum = doc.get_text()
        paragraphs = text_enum.get_paragraphs()

        for para in paragraphs:
            print("P--")
            portions = para.get_text_portions()
            for portion in portions:
                print(f' {portion.text_portion_type} = "{portion.get_string()}"')
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
    with BreakContext(
        Lo.Loader(connector=Lo.ConnectSocket(), opt=Lo.Options(verbose=args.verbose))
    ) as loader:
        fnm = cast(str, args.file_path)

        try:
            doc = WriteDoc.open_doc(fnm=fnm, loader=loader)
        except Exception as e:
            print(f"Could not open '{fnm}'")
            print(f"  {e}")
            # office will close and with statement is exited
            raise BreakContext.Break

        try:
            if visible:
                doc.set_visible()
            print_paras(doc)

        finally:
            doc.close_doc()

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
