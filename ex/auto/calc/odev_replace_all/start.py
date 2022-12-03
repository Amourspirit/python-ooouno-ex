#!/usr/bin/env python
import sys
import argparse

from ooodev.utils.file_io import FileIO
from replace_all import ReplaceAll


def args_add(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("search", help="One or more search terms", nargs="+")
    parser.add_argument(
        "-r",
        "--replace",
        help="Optional replace string",
        action="store",
        dest="replace",
        default=None,
    )
    parser.add_argument(
        "-a",
        "--search-all",
        help="Optional Set search to use search all over search iter.",
        action="store_true",
        dest="search_all",
        default=False,
    )
    parser.add_argument(
        "-o",
        "--out",
        help="Optional file path of output file",
        action="store",
        dest="out_file",
        default="",
    )


def main() -> int:
    # create parser to read terminal input
    parser = argparse.ArgumentParser(description="main")

    # add args to parser
    args_add(parser=parser)

    if len(sys.argv) == 1:
        sys.argv.append("cat")
    # read the current command line args
    args = parser.parse_args()

    rp = ReplaceAll(srch_strs=args.search, repl_str=args.replace, out_fnm=args.out_file, is_search_all=args.search_all)
    rp.main()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
