from __future__ import annotations
import sys
import argparse
from typing import cast
from pathlib import Path


from ooodev.write import Write
from ooodev.loader import Lo
from ooodev.utils.file_io import FileIO
from ooodev.utils.info import Info
from ooodev.wrapper.break_context import BreakContext

from com.sun.star.document import XDocumentInsertable
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
    parser.add_argument("docs", help="Append one or more documents", nargs="*")


def append_text_files(doc: XTextDocument, *args: str) -> None:
    cursor = Write.get_cursor(doc)
    for arg in args:
        try:
            cursor.gotoEnd(False)
            print(f"Appending {arg}")
            inserter = Lo.qi(XDocumentInsertable, cursor)
            if inserter is None:
                print("Document inserter could not be created")
            else:
                inserter.insertDocumentFromURL(FileIO.fnm_to_url(arg), ())
        except Exception as e:
            print(f"Could not append {arg} : {e}")


def main() -> int:
    # create parser to read terminal input
    parser = argparse.ArgumentParser(description="main")

    # add args to parser
    args_add(parser=parser)

    if len(sys.argv) <= 3:
        parser.print_help()
        return 0

    # read the current command line args
    args = parser.parse_args()
    # print(args)
    # return

    # Using Lo.Loader context manager warped by BreakContext load Office and connect via socket.
    # Context manager takes care of terminating instance when job is done.
    # see: https://python-ooo-dev-tools.readthedocs.io/en/latest/src/wrapper/break_context.html
    # see: https://python-ooo-dev-tools.readthedocs.io/en/latest/src/utils/lo.html#ooodev.utils.lo.Lo.Loader
    with BreakContext(Lo.Loader(Lo.ConnectSocket(headless=True))) as loader:

        fnm = cast(str, args.file_path)

        try:
            doc = Write.open_doc(fnm=fnm, loader=loader)
        except Exception as e:
            print(f"Could not open '{fnm}'")
            print(f"  {e}")
            # office will close and with statement is exited
            raise BreakContext.Break

        try:
            append_text_files(doc, *args.docs)
            pth = Path.cwd() / "tmp"
            pth.mkdir(exist_ok=True)
            doc_pth = pth / (Info.get_name(fnm) + "_APPENDED." + Info.get_ext(fnm))
            Lo.save_doc(doc, doc_pth)
        finally:
            Lo.close_doc(doc)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
