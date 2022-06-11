from __future__ import annotations
from ooodev.utils.lo import Lo
from ooodev.office.calc import Calc
from com.sun.star.sheet import XSpreadsheet

def do_cell_range(sheet: XSpreadsheet) -> None:
    Calc.highlight_range(sheet=sheet, headline="Range Data Example", range_name="A2:C24")
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

def create_array() -> None:
    loader = Lo.load_office(direct=True)
    doc = Calc.get_ss_doc(Lo.ThisComponent)
    sheet = Calc.get_active_sheet(doc=doc)
    do_cell_range(sheet=sheet)
    

def clear_range() -> None:
    loader = Lo.load_office(direct=True)
    doc = Calc.get_ss_doc(Lo.ThisComponent)
    sheet = Calc.get_active_sheet(doc=doc)
    flags = Calc.CellFlags.VALUE | Calc.CellFlags.STRING | Calc.CellFlags.FORMULA
    Calc.clear_cells(sheet=sheet, range_name="A2:C24", cell_flags=flags)
    Calc.remove_border(sheet=sheet,range_name="A2:C24")