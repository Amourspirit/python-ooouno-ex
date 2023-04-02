#!/usr/bin/env python
from __future__ import annotations
import uno
from com.sun.star.sheet import XSpreadsheet

from ooodev.dialog.msgbox import MsgBox, MessageBoxType, MessageBoxButtonsEnum, MessageBoxResultsEnum
from ooodev.utils.lo import Lo
from ooodev.utils.gui import GUI
from ooodev.office.calc import Calc
from ooodev.format.calc.direct.cell.borders import Borders, Side
from ooodev.utils.color import CommonColor


def do_cell_range(sheet: XSpreadsheet) -> None:
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
    Calc.set_array(values=vals, sheet=sheet, name="A3:C23")  # or just "A3"
    Calc.set_val("Total", sheet=sheet, cell_name="A24")
    Calc.set_val("=SUM(C4:C23)", sheet=sheet, cell_name="C24")

    # set Border around data and summary.
    bdr = Borders(border_side=Side(color=CommonColor.LIGHT_BLUE, width=2.85))
    Calc.set_style_range(sheet=sheet, range_name="A2:C24", styles=[bdr])


def main() -> int:
    _ = Lo.load_office(Lo.ConnectSocket())
    try:
        doc = Calc.create_doc()
        sheet = Calc.get_sheet(doc=doc, index=0)
        GUI.set_visible(is_visible=True, odoc=doc)
        Lo.delay(300)
        Calc.zoom(doc=doc, type=GUI.ZoomEnum.ZOOM_100_PERCENT)

        do_cell_range(sheet=sheet)
        Lo.delay(1_500)
        msg_result = MsgBox.msgbox(
            "Do you wish to close document?",
            "All done",
            boxtype=MessageBoxType.QUERYBOX,
            buttons=MessageBoxButtonsEnum.BUTTONS_YES_NO,
        )
        if msg_result == MessageBoxResultsEnum.YES:
            Lo.close_doc(doc=doc, deliver_ownership=True)
            Lo.close_office()
        else:
            print("Keeping document open")
    except Exception:
        Lo.close_office()
        raise
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
