#!/usr/bin/env python
"""
This module is not to be called directly
and is intended to be called from the projects main.
Such as: python -m main auto --process "ex/auto/calc/odev_add_range_data/start.py"
"""
from __future__ import annotations
from ooodev.utils.lo import Lo
from ooodev.utils.gui import GUI
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

def main() -> int:
    loader = Lo.load_office(Lo.ConnectSocket())
    doc = Calc.create_doc(loader=loader)
    sheet = Calc.get_sheet(doc=doc, index=0)
    GUI.set_visible(is_visible=True, odoc=doc)
    do_cell_range(sheet=sheet)
    return 0

if __name__ == "__main__":
    raise SystemExit(main())