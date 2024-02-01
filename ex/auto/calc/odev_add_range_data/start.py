#!/usr/bin/env python
from __future__ import annotations
import uno

from ooodev.dialog.msgbox import (
    MessageBoxType,
    MessageBoxButtonsEnum,
    MessageBoxResultsEnum,
)
from ooodev.utils.lo import Lo
from ooodev.calc import CalcDoc, CalcSheet, ZoomKind
from ooodev.format.calc.direct.cell.borders import Borders, Side
from ooodev.utils.color import CommonColor


def do_cell_range(sheet: CalcSheet) -> None:
    vals = (
        ("Name", "Fruit", "Quantity"),
        ("Alice", "Apples", 3),
        ("Alice", "Oranges", 7),
        ("Bob", "Apples", 3),
        ("Alice", "Apples", 9),
        ("Bob", "Apples", 5),
        ("Bob", "Oranges", 6),
        ("Alice", "Oranges", 3),
        ("Alice", "Apples", 8),
        ("Alice", "Oranges", 1),
        ("Bob", "Oranges", 2),
        ("Bob", "Oranges", 7),
        ("Bob", "Apples", 1),
        ("Alice", "Apples", 8),
        ("Alice", "Oranges", 8),
        ("Alice", "Apples", 7),
        ("Bob", "Apples", 1),
        ("Bob", "Oranges", 9),
        ("Bob", "Oranges", 3),
        ("Alice", "Oranges", 4),
        ("Alice", "Apples", 9),
    )
    sheet.set_array(values=vals, name="A3:C23")  # or just "A3"
    cell = sheet.get_cell(cell_name="A24")
    cell.set_val("Total")

    cell = sheet.get_cell(cell_name="C24")
    cell.set_val("=SUM(C4:C23)")

    # set Border around data and summary.
    bdr = Borders(border_side=Side(color=CommonColor.LIGHT_BLUE, width=2.85))
    rng = sheet.get_range(range_name="A2:C24")
    rng.set_style([bdr])


def main() -> int:
    _ = Lo.load_office(Lo.ConnectSocket())
    try:
        doc = CalcDoc.create_doc(visible=True)
        sheet = doc.sheets[0]
        # doc.set_visible()
        # wait for document to be ready
        Lo.delay(300)
        doc.zoom(ZoomKind.ZOOM_100_PERCENT)

        do_cell_range(sheet=sheet)
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
