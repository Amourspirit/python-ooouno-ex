from __future__ import annotations
import argparse
import sys
from pathlib import Path

from extract_graphics import ExtractGraphics


def args_add(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "-f",
        "--file",
        help="File path of input file",
        action="store",
        dest="file_path",
        required=True,
    )

    parser.add_argument(
        "-d",
        "--output_dir",
        help="Optional output Directory. Defaults to temporary dir sub folder.",
        action="store",
        dest="out_dir",
        default="",
    )


def main() -> int:
    if len(sys.argv) == 1:
        fnm = Path(__file__).parent / "data" / "build.odt"
        sys.argv.append("-f")
        sys.argv.append(str(fnm))

    # create parser to read terminal input
    parser = argparse.ArgumentParser(description="main")

    # add args to parser
    args_add(parser=parser)

    # read the current command line args
    args = parser.parse_args()
    eg = ExtractGraphics(fnm=args.file_path, out_dir=args.out_dir)
    eg.main()

    return 0


if __name__ == "__main__":
    SystemExit(main())
