import sys
from grouper import Grouper, CombineEllipseKind
import argparse


# region args parse
def args_add(parser: argparse.ArgumentParser) -> None:
    # usage for default start.py -k
    parser.add_argument(
        "-k",
        "--kind",
        const="combine",
        nargs="?",
        dest="kind",
        choices=[e.value for e in CombineEllipseKind],
        help="Kind of combining of ellipses combining to preform (default: %(default)s)",
    )
    parser.add_argument(
        "-o",
        "--overlap",
        help="Determines if ellipses are to overlap.",
        action="store_true",
        dest="overlap",
        default=False,
    )


# endregion args parse


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

    g = Grouper()
    kind = CombineEllipseKind(args.kind)
    g.overlap = args.overlap
    g.combine_kind = kind
    g.main()
    return 0


# endregion main()

if __name__ == "__main__":
    SystemExit(main())
