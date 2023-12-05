from __future__ import annotations

import uno
from com.sun.star.sheet import XDataPilotTable
from com.sun.star.sheet import XDataPilotTables
from com.sun.star.sheet import XSpreadsheet

from ooodev.dialog.msgbox import (
    MsgBox,
    MessageBoxType,
    MessageBoxButtonsEnum,
    MessageBoxResultsEnum,
)
from ooodev.calc import Calc, GeneralFunction
from ooodev.calc import CalcDoc, CalcSheet
from ooodev.utils.file_io import FileIO
from ooodev.utils.lo import Lo
from ooodev.utils.props import Props
from ooodev.utils.type_var import PathOrStr

from ooo.dyn.sheet.data_pilot_field_orientation import DataPilotFieldOrientation


class PivotTable2:
    def __init__(self, fnm: PathOrStr, out_fnm: PathOrStr) -> None:
        _ = FileIO.is_exist_file(fnm, True)
        self._fnm = FileIO.get_absolute_path(fnm)
        if out_fnm:
            out_file = FileIO.get_absolute_path(out_fnm)
            _ = FileIO.make_directory(out_file)
            self._out_fnm = out_file
        else:
            self._out_fnm = ""

    def main(self) -> None:
        loader = Lo.load_office(Lo.ConnectSocket())

        try:
            doc = CalcDoc(Calc.open_doc(fnm=self._fnm, loader=loader))

            doc.set_visible()

            sheet = doc.get_sheet(0)
            dp_sheet = doc.insert_sheet(name="Pivot Table", idx=1)

            self._create_pivot_table(sheet=sheet, dp_sheet=dp_sheet)
            dp_sheet.set_active()

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

    def _create_pivot_table(
        self, sheet: CalcSheet, dp_sheet: CalcSheet
    ) -> XDataPilotTable | None:
        cell_range = sheet.find_used_range()
        print(f"The used area is: { cell_range.get_range_str()}")
        print()

        dp_tables = sheet.get_pilot_tables()
        dp_desc = dp_tables.createDataPilotDescriptor()
        dp_desc.setSourceRange(cell_range.get_address())

        # XIndexAccess fields = dpDesc.getDataPilotFields();
        fields = dp_desc.getHiddenFields()
        field_names = Lo.get_container_names(con=fields)
        print(f"Field Names ({len(field_names)}):")
        for name in field_names:
            print(f"  {name}")

        # properties defined in DataPilotField
        # set page field
        props = Lo.find_container_props(con=fields, nm="Date")
        Props.set(props, Orientation=DataPilotFieldOrientation.PAGE)

        # set column field
        props = Lo.find_container_props(con=fields, nm="Store")
        Props.set(props, Orientation=DataPilotFieldOrientation.COLUMN)

        # set 1st row field
        props = Lo.find_container_props(con=fields, nm="Book")
        Props.set(props, Orientation=DataPilotFieldOrientation.ROW)

        # set data field, calculating the sum
        props = Lo.find_container_props(con=fields, nm="Units Sold")
        Props.set(props, Orientation=DataPilotFieldOrientation.DATA)
        Props.set(props, Function=GeneralFunction.SUM)

        # place onto sheet
        dest_addr = dp_sheet.get_cell_address(cell_name="A1")
        dp_tables.insertNewByName("PivotTableExample", dest_addr, dp_desc)
        dp_sheet.set_col_width(width=60, idx=0)
        # A column; in mm

        # Usually the table is not fully updated. The cells are often
        # drawn with #VALUE! contents (?).

        # This can be fixed by explicitly refreshing the table, but it has to
        # be accessed via the sheet or the tables container is considered
        # empty, and the table is not found.

        dp_tables2 = dp_sheet.get_pilot_tables()
        return self._refresh_table(dp_tables=dp_tables2, table_name="PivotTableExample")

    def _refresh_table(
        self, dp_tables: XDataPilotTables, table_name: str
    ) -> XDataPilotTable | None:
        nms = dp_tables.getElementNames()
        print(f"No. of DP tables: {len(nms)}")
        for nm in nms:
            print(f"  {nm}")

        dp_table = Calc.get_pilot_table(dp_tables=dp_tables, name=table_name)
        if dp_table is not None:
            dp_table.refresh()
        return dp_table
