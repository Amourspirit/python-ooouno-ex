import sys
from pathlib import Path
from ooodev.utils.file_io import FileIO
from slides_info import SlidesInfo


# region main()
def main() -> int:
    if len(sys.argv) > 1:
        fnm = sys.argv[1]
        FileIO.is_exist_file(fnm, True)
    else:
        fnm = Path(__file__).parent / "data" / "algs.odp"

    slides_info = SlidesInfo(fnm)
    slides_info.main()
    return 0


# endregion main()

if __name__ == "__main__":
    SystemExit(main())
