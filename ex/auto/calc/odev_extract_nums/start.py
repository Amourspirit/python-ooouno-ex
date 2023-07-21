from pathlib import Path
from extract_nums import ExtractNums


def main() -> int:
    fnm = Path(__file__).parent / "data" / "small_totals.ods"

    en = ExtractNums(fnm)
    en.main()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
