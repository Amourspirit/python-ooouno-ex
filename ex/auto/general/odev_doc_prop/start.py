#!/usr/bin/env python
# coding: utf-8
# region Imports
from __future__ import annotations
import argparse


from ooodev.utils.lo import Lo
from ooodev.utils.info import Info

# endregion Imports

# region args
def args_add(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "-d",
        "--doc",
        help="Show Doc properties",
        action="store",
        dest="fnm_doc",
        required=True,
    )


# endregion args

# region Main
def main() -> int:
    # create parser to read terminal input
    parser = argparse.ArgumentParser(description="main")

    # add args to parser
    args_add(parser=parser)

    # read the current command line args
    args = parser.parse_args()

    with Lo.Loader(Lo.ConnectSocket(headless=True)) as loader:
        doc = Lo.open_doc(fnm=args.fnm_doc, loader=loader)
        Info.print_doc_properties(doc)
        Info.set_doc_props(doc, "Example", "Examples", "Specific User")
        
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

# endregion main
