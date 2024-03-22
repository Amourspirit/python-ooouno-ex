from ooodev.calc import CalcDoc
from tab_dialog import Tabs


def show_tabs(*args) -> None:
    doc = CalcDoc.from_current_doc()
    tabs = Tabs(doc=doc)
    tabs.show()
