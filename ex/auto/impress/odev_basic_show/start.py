import sys
from ooodev.utils.file_io import FileIO
from pathlib import Path
from basic_show import BasicShow

# region main()
def main() -> int:
    if len(sys.argv) > 1:
        fnm = sys.argv[1]
        FileIO.is_exist_file(fnm, True)
    else:
        fnm = Path(__file__).parent / "data" / "algs.ppt"

    basic_show = BasicShow(fnm)
    basic_show.main()
    return 0


# endregion main()

if __name__ == "__main__":
    SystemExit(main())
