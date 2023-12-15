import sys
from pathlib import Path
from extract_text import ExtractText


# region main()
def main() -> int:
    if len(sys.argv) == 2:
        fnm = sys.argv[1]
    else:
        fnm = Path(__file__).parent / "data" / "algs.odp"
    et = ExtractText(fnm)
    et.main()
    return 0


# endregion main()

if __name__ == "__main__":
    SystemExit(main())
