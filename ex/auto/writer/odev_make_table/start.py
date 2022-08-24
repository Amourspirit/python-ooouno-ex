#!/usr/bin/env python
# coding: utf-8
import csv
from pathlib import Path
from typing import List

from ooodev.utils.lo import Lo
from ooodev.office.write import Write
from ooodev.utils.gui import GUI
from ooodev.utils.file_io import FileIO
from ooodev.utils.date_time_util import DateUtil


def read_table(fnm: Path) -> List[list]:
    # get a 2D Table with the the first row containing column names.
    results = []
    with open(fnm, "r", newline="") as csv_file:
        csv_reader = csv.reader(csv_file, delimiter="\t")
        line_count = 0
        for row in csv_reader:
            # each row is a list of values
            if line_count in (0, 1, 2, 4):  # skip non-csv lines
                line_count += 1
                continue
            # first row will be column names
            results.append(row)
            line_count += 1
    return results


def main() -> int:

    fnm = FileIO.get_absolute_path("../../../../resources/txt/bondMovies.txt")  # source csv file
    if not fnm.exists():
        fnm = FileIO.get_absolute_path("resources/txt/bondMovies.txt")
    if not fnm.exists():
        print("resource image 'bondMovies.txt' not found.")
        print("Unable to continue.")
        return 1

    tbl_data = read_table(fnm)

    delay = 2_000  # delay so users can see changes.

    with Lo.Loader(Lo.ConnectSocket()) as loader:

        doc = Write.create_doc(loader=loader)

        try:
            GUI.set_visible(is_visible=True, odoc=doc)

            cursor = Write.get_cursor(doc)

            Write.append_para(cursor, "Table of Bond Movies")
            Write.style_prev_paragraph(cursor, "Heading 1")
            Write.append_para(cursor, 'The following table comes form "bondMovies.txt"\n')

            # Lock display updating for faster writing of table into document.
            with Lo.ControllerLock():
                Write.add_table(cursor=cursor, table_data=tbl_data)
                Write.end_paragraph(cursor)

            Lo.delay(delay)
            Write.append(cursor, f"Timestamp: {DateUtil.time_stamp()}")
            Lo.delay(delay)
            Lo.save_doc(doc, "table.odt")

        finally:
            Lo.close_doc(doc)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
