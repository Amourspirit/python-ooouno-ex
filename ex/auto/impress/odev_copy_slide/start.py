import sys
from pathlib import Path
from copy_slide import CopySlide


# region main()
def main() -> int:
    if len(sys.argv) == 4:
        p = Path(sys.argv[1])
        from_idx = int(sys.argv[2])
        to_idx = int(sys.argv[3])

    else:
        from_idx = 2
        to_idx = 4
        p = Path(__file__).parent / "data" / "algs.odp"
    # slide indexes are zero based indexes.
    cs = CopySlide(fnm=p, from_idx=from_idx, to_idx=to_idx)
    cs.main()
    return 0


# endregion main()

if __name__ == "__main__":
    SystemExit(main())
