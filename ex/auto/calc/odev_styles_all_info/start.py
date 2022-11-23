#!/usr/bin/env python
import sys
import argparse
from ooodev.utils.file_io import FileIO
from styles_all_info import StylesAllInfo

def args_add(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "-f",
        "--file",
        help="File path of input file",
        action="store",
        dest="file_path",
        required=False,
    )
    parser.add_argument(
        "-r",
        "--report-props",
        help="Optionally show long report",
        action="store_true",
        dest="rpt",
        default=False
    )

def main() -> int:
    # create parser to read terminal input
    parser = argparse.ArgumentParser(description="main")
    
    # add args to parser
    args_add(parser=parser)

    if len(sys.argv) == 1:
        fnm = "resources/ods/totals.ods"
        if not FileIO.is_exist_file(fnm):
            fnm = "../../../../resources/ods/totals.ods"
        sys.argv.append('-f')
        sys.argv.append(fnm)

    args = parser.parse_args()

    sinfo = StylesAllInfo(args.file_path, args.rpt)
    sinfo.main()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
