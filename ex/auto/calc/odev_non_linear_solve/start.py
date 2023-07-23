import argparse

from solver1 import Solver1
from solver2 import Solver2


def args_add(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "-s",
        "--solver",
        nargs="?",
        dest="ex",
        choices=[1, 2],
        default=1,
        type=int,
        help="Which of solver example to run (default: %(default)s)",
    )
    parser.add_argument("-v", "--verbose", help="Verbose output", action="store_true", dest="verbose", default=False)


def main() -> int:
    # create parser to read terminal input
    parser = argparse.ArgumentParser(description="main")

    # add args to parser
    args_add(parser=parser)

    # read the current command line args
    args = parser.parse_args()

    if args.ex == 1:
        Solver1.main(args.verbose)
    else:
        Solver2.main(args.verbose)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
