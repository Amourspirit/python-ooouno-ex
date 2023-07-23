from __future__ import annotations
from pathlib import Path
from append_slides import AppendSlides
import sys


def main() -> int:
    fnm_lst = []
    if len(sys.argv) > 1:
        fnm_lst.extend(sys.argv[1:])
    else:

        files = ("algs.odp", "points.odp")
        dir_path = Path(__file__).parent / "data"
        for file in files:
            fnm_lst.append(Path(dir_path, file))

    appended = AppendSlides(*fnm_lst)
    appended.main()

    return 0


if __name__ == "__main__":
    SystemExit(main())
