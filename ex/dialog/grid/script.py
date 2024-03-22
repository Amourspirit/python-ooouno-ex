from ooodev.calc import CalcDoc
from grid_ex import GridEx


def show_grid(*args) -> None:
    doc = CalcDoc.from_current_doc()
    grid_ex = GridEx(doc=doc)
    grid_ex.show()
