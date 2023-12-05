from __future__ import annotations

import uno
from com.sun.star.util import XSortable

from ooodev.dialog.msgbox import (
    MsgBox,
    MessageBoxType,
    MessageBoxButtonsEnum,
    MessageBoxResultsEnum,
)
from ooodev.calc import Calc
from ooodev.calc import CalcDoc
from ooodev.utils.file_io import FileIO
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.utils.props import Props
from ooodev.utils.type_var import PathOrStr

from ooo.dyn.table.table_sort_field import TableSortField


class DataSort:
    def __init__(self, out_fnm: PathOrStr, **kwargs) -> None:
        if out_fnm:
            out_f = FileIO.get_absolute_path(out_fnm)
            _ = FileIO.make_directory(out_f)
            self._out_fnm = out_f
        else:
            self._out_fnm = ""

    def main(self) -> None:
        loader = Lo.load_office(Lo.ConnectSocket())

        try:
            doc = CalcDoc(Calc.create_doc(loader))

            doc.set_visible()

            sheet = doc.get_sheet(0)

            # create the table that needs sorting
            vals = (
                ("Level", "Code", "No.", "Team", "Name"),
                ("BS", 20, 4, "B", "Elle"),
                ("BS", 20, 6, "C", "Sweet"),
                ("BS", 20, 2, "A", "Chcomic"),
                ("CS", 30, 5, "A", "Ally"),
                ("MS", 10, 1, "A", "Joker"),
                ("MS", 10, 3, "B", "Kevin"),
                ("CS", 30, 7, "C", "Tom"),
            )
            sheet.set_array(values=vals, name="A1:E8")  # or just "A1"

            # 1. obtain an XSortable interface for the cell range
            source_range = sheet.get_range(range_name="A1:E8")
            x_sort = source_range.qi(XSortable, True)

            # 2. specify the sorting criteria as a TableSortField array
            sort_fields = (self._make_sort_asc(1, True), self._make_sort_asc(2, True))

            # 3. define a sort descriptor
            props = Props.make_props(
                SortFields=Props.any(*sort_fields), ContainsHeader=True
            )

            Lo.wait(2_000)  # wait so user can see original before it is sorted
            # 4. do the sort
            print("Sorting...")
            x_sort.sort(props)

            if self._out_fnm:
                doc.save_doc(fnm=self._out_fnm)

            msg_result = MsgBox.msgbox(
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

    def _make_sort_asc(self, index: int, is_ascending: bool) -> TableSortField:
        return TableSortField(
            Field=index, IsAscending=is_ascending, IsCaseSensitive=False
        )
