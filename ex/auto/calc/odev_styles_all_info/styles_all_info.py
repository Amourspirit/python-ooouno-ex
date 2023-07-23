from __future__ import annotations

import uno
from com.sun.star.sheet import XSpreadsheetDocument

from ooodev.office.calc import Calc
from ooodev.utils.file_io import FileIO
from ooodev.utils.info import Info
from ooodev.utils.lo import Lo
from ooodev.utils.props import Props
from ooodev.utils.type_var import PathOrStr


class StylesAllInfo:
    def __init__(self, fnm: PathOrStr, rpt_cell_styles: bool) -> None:
        _ = FileIO.is_exist_file(fnm, True)
        self._fnm = FileIO.get_absolute_path(fnm)
        self._rpt_cell_styles = rpt_cell_styles

    def main(self) -> None:
        with Lo.Loader(Lo.ConnectSocket(headless=True)) as loader:
            doc = Calc.open_doc(fnm=self._fnm, loader=loader)
            try:
                # get all the style families for this document
                style_families = Info.get_style_family_names(doc)
                print(f"Style Family Names ({len(style_families)})")
                for style_family in style_families:
                    print(f"  {style_family}")
                print()

                # list all the style names for each style family
                for i, style_family in enumerate(style_families):
                    print(f'{i + 1}. "{style_family}" Family Styles:')
                    style_names = Info.get_style_names(doc=doc, family_style_name=style_family)
                    Lo.print_names(names=style_names)

                if self._rpt_cell_styles:
                    print()
                    self._report_cell_styles(doc)

            except Exception:
                raise

            finally:
                Lo.close_doc(doc=doc, deliver_ownership=True)

    def _report_cell_styles(self, doc: XSpreadsheetDocument) -> None:
        Props.show_props(
            "CellStyles Default", Info.get_style_props(doc=doc, family_style_name="CellStyles", prop_set_nm="Default")
        )

        Props.show_props(
            "PageStyles Default", Info.get_style_props(doc=doc, family_style_name="PageStyles", prop_set_nm="Default")
        )
