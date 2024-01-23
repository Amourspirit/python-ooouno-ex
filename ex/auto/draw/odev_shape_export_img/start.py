from __future__ import annotations
from pathlib import Path
import uno
from bezier_builder import BezierBuilder
import sys


def main() -> int:
    arg = "png"
    if len(sys.argv) > 1:
        try:
            arg = str(sys.argv[1])
        except ValueError:
            arg = "png"
    tmp = Path.cwd() / "tmp"
    tmp.mkdir(parents=True, exist_ok=True)
    if arg == "png":
        fnm = tmp / "closed_bezier.png"
    else:
        fnm = tmp / "closed_bezier.jpg"
    builder = BezierBuilder(fnm)
    builder.export(200)
    return 0


if __name__ == "__main__":
    SystemExit(main())
