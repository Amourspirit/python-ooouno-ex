from __future__ import annotations
import sys
from pathlib import Path
from chart_2_views import Chart2View, ChartKind
from ooodev.utils.file_io import FileIO
import argparse


def args_add(parser: argparse.ArgumentParser) -> None:
    # usage for default start.py -k
    parser.add_argument(
        "-k",
        "--kind",
        const="happy_stock",
        nargs="?",
        dest="kind",
        choices=[e.value for e in ChartKind],
        help="Kind of chart to display (default: %(default)s)",
    )


# region main()
def main() -> int:
    if len(sys.argv) == 1:
        sys.argv.append("-k")
    # create parser to read terminal input
    parser = argparse.ArgumentParser(description="main")

    # add args to parser
    args_add(parser=parser)

    # read the current command line args
    args = parser.parse_args()

    fnm = Path(__file__).parent / "data" / "chartsData.ods"

    kind = ChartKind(args.kind)

    cv = Chart2View(data_fnm=fnm, chart_kind=kind)
    cv.main()
    return 0


# endregion main()

if __name__ == "__main__":
    SystemExit(main())
