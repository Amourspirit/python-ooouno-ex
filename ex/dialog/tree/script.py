from ooodev.loader import Lo
from tab_dialog import Tabs

# import the WriteDoc class just to ensure that it is included in the full script compile.
from ooodev.write import WriteDoc


def show_tabs(*args) -> None:
    _ = Lo.current_doc
    tabs = Tabs()
    tabs.show()
