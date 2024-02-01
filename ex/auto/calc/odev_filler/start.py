#!/usr/bin/env python
import argparse
from filler import Filler


def args_add(parser: argparse.ArgumentParser) -> None:
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

    # read the current command line args
    args = parser.parse_args()

    fl = Filler(out_fnm=args.out_file)
    fl.main()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
