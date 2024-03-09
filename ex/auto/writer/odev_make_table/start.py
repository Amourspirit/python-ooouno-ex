from __future__ import annotations
import csv
from pathlib import Path
from typing import List
import uno
from ooo.dyn.style.paragraph_adjust import ParagraphAdjust

from ooodev.dialog.msgbox import (
    MessageBoxType,
    MessageBoxButtonsEnum,
    MessageBoxResultsEnum,
)
from ooodev.loader import Lo
from ooodev.format.writer.style.para import Para as ParaStyle
from ooodev.format.inner.direct.structs.side import LineSize
from ooodev.utils.date_time_util import DateUtil
from ooodev.utils.color import StandardColor
from ooodev.write import WriteDoc


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
    fnm = Path(__file__).parent / "data" / "bondMovies.txt"  # source csv file

    tbl_data = read_table(fnm)

    delay = 2_000  # delay so users can see changes.

    loader = Lo.load_office(Lo.ConnectSocket())

    try:
        doc = WriteDoc.create_doc(loader=loader, visible=True)

        cursor = doc.get_cursor()
        cursor.append_para("Table of Bond Movies", styles=[ParaStyle().h1])
        cursor.append_para('The following table comes form "bondMovies.txt"\n')

        # Lock display updating for faster writing of table into document.
        with doc:
            tbl = cursor.add_table(table_data=tbl_data, first_row_header=True)
            cursor.end_paragraph()

            tbl.style_direct.style_borders_side(
                color=StandardColor.BLUE_DARK1, width=LineSize.MEDIUM
            )
            tbl.table_column_separators[0].position += 1000  # make first column wider
            # set the table border line color
            tbl.table_border.horizontal_line.color = StandardColor.BLUE_LIGHT1
            tbl.table_border.vertical_line.color = StandardColor.BLUE_LIGHT1

            # iterate over the rows and set the background color
            for i, row in enumerate(tbl.rows):
                if i == 0:
                    row.back_color = StandardColor.BLUE
                elif i % 2 == 0:
                    row.back_color = StandardColor.GRAY_LIGHT2
                else:
                    row.back_color = StandardColor.GRAY_LIGHT4

            header_row = tbl.rows[0]
            # style header row text to be white and bold
            for i, cell in enumerate(header_row):
                cell.style_direct.style_font_general(color=StandardColor.WHITE, b=True)

            # make the first column blue.
            col1 = tbl.columns[0]
            for i, cell in enumerate(col1):
                if i > 0:
                    cell.style_direct.style_font_general(color=StandardColor.BLUE_DARK3)

            # right align the second column and set the style to general numbers.
            col2 = tbl.columns[1]
            for i, cell in enumerate(col2):
                if i > 0:
                    cell.style_direct.style_numbers_general()
                    cell.style_direct.style_alignment(align=ParagraphAdjust.RIGHT)

        Lo.delay(delay)
        # Append a timestamp to the document.
        cursor.append(f"Timestamp: {DateUtil.time_stamp()}")
        Lo.delay(delay)
        msg_result = doc.msgbox(
            "Do you wish to save document?",
            "Save",
            boxtype=MessageBoxType.QUERYBOX,
            buttons=MessageBoxButtonsEnum.BUTTONS_YES_NO,
        )
        if msg_result == MessageBoxResultsEnum.YES:
            tmp = Path.cwd() / "tmp"
            tmp.mkdir(exist_ok=True)
            doc.save_doc(tmp / "table.odt")

        msg_result = doc.msgbox(
            "Do you wish to close document?",
            "All done",
            boxtype=MessageBoxType.QUERYBOX,
            buttons=MessageBoxButtonsEnum.BUTTONS_YES_NO,
        )
        if msg_result == MessageBoxResultsEnum.YES:
            doc.close_doc()
            Lo.close_office()
        else:
            print("Keeping document open")
    except Exception:
        Lo.close_office()
        raise

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
