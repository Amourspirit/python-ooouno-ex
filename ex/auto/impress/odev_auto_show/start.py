import sys
from pathlib import Path
from ooodev.utils.file_io import FileIO
from auto_show import AutoShow  # , FadeEffect


# region main()
def main() -> int:
    if len(sys.argv) > 1:
        fnm = sys.argv[1]
        FileIO.is_exist_file(fnm, True)
    else:
        fnm = Path(__file__).parent / "data" / "algs.ppt"

    auto_show = AutoShow(fnm)
    # auto_show.duration = 2
    # auto_show.fade_effect = FadeEffect.MOVE_FROM_LEFT
    # auto_show.end_delay = 3
    auto_show.main()
    return 0


# endregion main()

if __name__ == "__main__":
    SystemExit(main())
