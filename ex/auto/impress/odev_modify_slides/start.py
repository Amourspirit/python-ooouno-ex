from __future__ import annotations
from pathlib import Path
from ooodev.utils.file_io import FileIO
from modify_slides import ModifySlides
import sys


def main() -> int:
    data_dir = Path(__file__).parent / "data"
    im_fnm = data_dir / "questions.png"

    if len(sys.argv) > 1:
        fnm = sys.argv[1]
        _ = FileIO.is_exist_file(fnm, True)
    else:
        fnm = data_dir / "algsSmall.ppt"

    modify = ModifySlides(fnm=fnm, im_fnm=im_fnm)
    modify.main()

    return 0


if __name__ == "__main__":
    SystemExit(main())
