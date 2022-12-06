from __future__ import annotations

import uno
from com.sun.star.sheet import XDataPilotTable
from com.sun.star.sheet import XDataPilotTables
from com.sun.star.sheet import XSpreadsheet

from ooodev.dialog.msgbox import MsgBox, MessageBoxType, MessageBoxButtonsEnum, MessageBoxResultsEnum
from ooodev.office.calc import Calc, GeneralFunction
from ooodev.utils.file_io import FileIO
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.utils.props import Props
from ooodev.utils.type_var import PathOrStr

from ooo.dyn.sheet.data_pilot_field_orientation import DataPilotFieldOrientation


class PivotTable2:
    def __init__(self, fnm: PathOrStr, out_fnm: PathOrStr) -> None:
        _ = FileIO.is_exist_file(fnm, True)
        self._fnm = FileIO.get_absolute_path(fnm)
        if out_fnm:
            outf = FileIO.get_absolute_path(out_fnm)
            _ = FileIO.make_directory(outf)
            self._out_fnm = outf
        else:
            self._out_fnm = ""

    def main(self) -> None:
        loader = Lo.load_office(Lo.ConnectSocket())

        try:
            doc = Calc.open_doc(fnm=self._fnm, loader=loader)

            GUI.set_visible(is_visible=True, odoc=doc)

            sheet = Calc.get_sheet(doc=doc)
            dp_sheet = Calc.insert_sheet(doc=doc, name="Pivot Table", idx=1)

            self._create_pivot_table(sheet=sheet, dp_sheet=dp_sheet)
            Calc.set_active_sheet(doc=doc, sheet=dp_sheet)

            if self._out_fnm:
                Lo.save_doc(doc=doc, fnm=self._out_fnm)

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

    def _create_pivot_table(self, sheet: XSpreadsheet, dp_sheet: XSpreadsheet) -> XDataPilotTable | None:
        cell_range = Calc.find_used_range(sheet)
        print(f"The used area is: { Calc.get_range_str(cell_range)}")
        print()

        dp_tables = Calc.get_pilot_tables(sheet)
        dp_desc = dp_tables.createDataPilotDescriptor()
        dp_desc.setSourceRange(Calc.get_address(cell_range))

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
        dest_addr = Calc.get_cell_address(sheet=dp_sheet, cell_name="A1")
        dp_tables.insertNewByName("PivotTableExample", dest_addr, dp_desc)
        Calc.set_col_width(sheet=dp_sheet, width=60, idx=0)
        # A column; in mm

        # Usually the table is not fully updated. The cells are often
        # drawn with #VALUE! contents (?).

        # This can be fixed by explicitly refreshing the table, but it has to
        # be accessed via the sheet or the tables container is considered
        # empty, and the table is not found.

        dp_tables2 = Calc.get_pilot_tables(sheet=dp_sheet)
        return self._refresh_table(dp_tables=dp_tables2, table_name="PivotTableExample")

    def _refresh_table(self, dp_tables: XDataPilotTables, table_name: str) -> XDataPilotTable | None:
        nms = dp_tables.getElementNames()
        print(f"No. of DP tables: {len(nms)}")
        for nm in nms:
            print(f"  {nm}")

        dp_table = Calc.get_pilot_table(dp_tables=dp_tables, name=table_name)
        if dp_table is not None:
            dp_table.refresh()
        return dp_table
