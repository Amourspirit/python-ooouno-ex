#!/usr/bin/env python
import argparse
from linear_solve import LinearSolve


def args_add(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("-v", "--verbose", help="Verbose output", action="store_true", dest="verbose", default=False)


def main() -> int:
    parser = argparse.ArgumentParser(description="main")
    args_add(parser=parser)
    args = parser.parse_args()

    LinearSolve.main(args.verbose)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
