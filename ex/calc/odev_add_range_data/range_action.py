from __future__ import annotations
from ooo.dyn.sheet.cell_flags import CellFlagsEnum
from ooodev.calc import CalcDoc, CalcSheet
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
    cell = sheet["A24"]
    cell.value = "Total"

    cell = sheet["C24"]
    cell.value = "=SUM(C4:C23)"

    # set Border around data and summary.
    bdr = Borders(border_side=Side(color=CommonColor.LIGHT_BLUE, width=2.85))
    rng = sheet.get_range(range_name="A3:C24")
    rng.apply_styles(bdr)
    # Calc.set_style_range(sheet=sheet.component, range_name="A3:C24", styles=[bdr])


def create_array() -> None:
    # get access to current Calc Document
    doc = CalcDoc.from_current_doc()

    # get access to current spreadsheet
    sheet = doc.get_active_sheet()

    # insert the array of data
    do_cell_range(sheet=sheet)


def clear_range() -> None:
    # get access to current Calc Document
    doc = CalcDoc.from_current_doc()

    # get access to current spreadsheet
    sheet = doc.get_active_sheet()

    # create the flags that let Calc know what kind or data to remove from cells
    flags = CellFlagsEnum.VALUE | CellFlagsEnum.STRING | CellFlagsEnum.FORMULA

    # clears the cells in a given range
    sheet.clear_cells(range_name="A2:C24", cell_flags=flags)

    # remove the border from the range.
    # OooDev >= 0.25.2
    sheet.get_range(range_name="A2:C24").remove_border()
    # OooDev <= 0.25.1
    # _ = Calc.remove_border(sheet=sheet.component, range_name="A2:C24")
