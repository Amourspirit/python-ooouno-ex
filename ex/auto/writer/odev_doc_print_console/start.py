from __future__ import annotations
import sys
import argparse
from pathlib import Path
from typing import cast

from ooodev.write import Write
from ooodev.loader import Lo
from ooodev.utils.info import Info
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


def main() -> int:
    # create parser to read terminal input
    parser = argparse.ArgumentParser(description="main")

    # add args to parser
    args_add(parser=parser)

    if len(sys.argv) == 1:
        fnm = Path(__file__).parent / "data" / "cicero_dummy.odt"
        sys.argv.append("-f")
        sys.argv.append(str(fnm))

    # read the current command line args
    args = parser.parse_args()

    # Using Lo.Loader context manager warped by BreakContext load Office and connect via socket.
    # Context manager takes care of terminating instance when job is done.
    # Note the use of the headless flag. Not using GUI for process.
    # see: https://python-ooo-dev-tools.readthedocs.io/en/latest/src/wrapper/break_context.html
    # see: https://python-ooo-dev-tools.readthedocs.io/en/latest/src/utils/lo.html#ooodev.utils.lo.Lo.Loader
    with BreakContext(Lo.Loader(connector=Lo.ConnectSocket(headless=True))) as loader:

        fnm = cast(str, args.file_path)

        try:
            doc = Lo.open_doc(fnm=fnm, loader=loader)
        except Exception:
            print(f"Could not open '{fnm}'")
            raise BreakContext.Break

        if Info.is_doc_type(obj=doc, doc_type=Lo.Service.WRITER):
            text_doc = Write.get_text_doc(doc=doc)
            cursor = Write.get_cursor(text_doc)
            text = Write.get_all_text(cursor)
            print("Text Content".center(50, "-"))
            print(text)
            print("-" * 50)
        else:
            print("Extraction unsupported for this doc type")
        Lo.close_doc(doc)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
