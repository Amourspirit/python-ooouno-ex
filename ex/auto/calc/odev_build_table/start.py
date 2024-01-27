#
# on wayland (some versions of Linux)
# may get error:
#    (soffice:67106): Gdk-WARNING **: 02:35:12.168: XSetErrorHandler() called with a GDK error trap pushed. Don't do that.
# This seems to be a Wayland/Java compatibility issues.
# see: http://www.babelsoft.net/forum/viewtopic.php?t=24545

import argparse
import uno
from pathlib import Path

from build_table import BuildTable


def args_add(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "-o",
        "--out",
        help="Optional file path of output file",
        action="store",
        dest="out_file",
        default="",
    )
    parser.add_argument(
        "-p",
        "--add-pic",
        help="Optional add picture to sheet",
        action="store_true",
        dest="add_pic",
        default=False,
    )
    parser.add_argument(
        "-c",
        "--add-chart",
        help="Optional add chart to sheet",
        action="store_true",
        dest="add_chart",
        default=False,
    )
    parser.add_argument(
        "-s",
        "--no-style",
        help="Optional add chart to sheet",
        action="store_false",
        dest="add_style",
        default=True,
    )


def main() -> int:
    # create parser to read terminal input
    parser = argparse.ArgumentParser(description="main")

    # add args to parser
    args_add(parser=parser)

    # if len(sys.argv) <= 1:
    #     parser.print_help()
    #     return 0

    # read the current command line args
    args = parser.parse_args()

    im_fnm = Path(__file__).parent / "image" / "skinner.png"

    bt = BuildTable(
        im_fnm=im_fnm,
        out_fnm=args.out_file,
        add_pic=args.add_pic,
        add_chart=args.add_chart,
        add_style=args.add_style,
    )
    bt.main()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
