#!/usr/bin/env python
from __future__ import annotations
import uno

from ooodev.dialog.msgbox import (
    MessageBoxType,
    MessageBoxButtonsEnum,
    MessageBoxResultsEnum,
)
from ooodev.loader import Lo
from ooodev.calc import CalcDoc, CalcSheet, ZoomKind
from ooodev.utils.color import CommonColor


def set_cell_data(sheet: CalcSheet) -> None:
    vals = (
        ("Hyperlinks",),
        ("https://ask.libreoffice.org/t/how-to-convert-links-into-hyperlinks-in-bulk-in-calc/102448",),
        ("https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/calc/odev_garlic_secrets",),
        ("https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/calc/odev_cell_texts",),
        ("https://python-ooo-dev-tools.readthedocs.io/en/latest/src/utils/data_type/range_obj.html",),
        ("https://python-ooo-dev-tools.readthedocs.io/en/latest/index.html",),
        ("https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part4/index.html",),
    )
    sheet.set_array(values=vals, name="A1")
    # bold the header
    _ = sheet["A1"].style_font_general(b=True)
    col = sheet.get_col_range(0)
    col.optimal_width = True

def create_hyperlinks() -> None:
    # get access to current Calc Document
    doc = CalcDoc.from_current_doc()

    # get access to first spreadsheet
    sheet = doc.sheets[0]

    # insert the array of data
    set_cell_data(sheet=sheet)
    convert_to_hyperlinks(sheet=sheet)

def convert_to_hyperlinks(sheet: CalcSheet) -> None:
    used_rng = sheet.find_used_range_obj()
    data = sheet.get_array(range_obj=used_rng)
    row_count = 0
    for row in data:
        for i, cell_data in enumerate(row):
            if cell_data.startswith("http"):
                cell = sheet[(i, row_count)]
                cell.value = ""
                cursor = cell.create_text_cursor()
                cursor.add_hyperlink(
                    label=cell_data,
                    url_str=cell_data,
                )
        row_count += 1

def main() -> int:
    _ = Lo.load_office(Lo.ConnectSocket())
    try:
        doc = CalcDoc.create_doc(visible=True)
        # wait for document to be ready
        Lo.delay(300)
        doc.zoom(ZoomKind.ZOOM_100_PERCENT)

        create_hyperlinks()
        Lo.delay(1_500)
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
