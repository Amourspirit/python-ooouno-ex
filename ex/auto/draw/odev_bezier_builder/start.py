from __future__ import annotations
import uno
from pathlib import Path
from ooodev.utils.file_io import FileIO
from bezier_builder import BezierBuilder
import sys


def main() -> int:
    arg = 2
    if len(sys.argv) > 1:
        try:
            arg = int(sys.argv[1])
            if arg < 0 or arg > 3:
                arg = 2
        except ValueError:
            arg = 2
    # file name: bpts0.txt or bpts1.txt or bpts2.txt or bpts3.txt
    file_name = f"bpts{arg}.txt"
    fnm = Path(__file__).parent / "data" /file_name

    builder = BezierBuilder(fnm_point=fnm)
    builder.show()
    return 0


if __name__ == "__main__":
    SystemExit(main())
