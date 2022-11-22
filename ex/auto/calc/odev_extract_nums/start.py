#!/usr/bin/env python
from ooodev.utils.file_io import FileIO
from extract_nums import ExtractNums



def main() -> int:
    fnm = "resources/ods/small_totals.ods"
    if not FileIO.is_exist_file(fnm):
        fnm = "../../../../resources/ods/small_totals.ods"

    en = ExtractNums(fnm)
    en.main()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
