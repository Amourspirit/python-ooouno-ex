from __future__ import annotations
from typing import cast


import uno
from com.sun.star.sheet import XCellRangesQuery

from ooodev.dialog.msgbox import MsgBox, MessageBoxType, MessageBoxButtonsEnum, MessageBoxResultsEnum
from ooodev.office.calc import Calc
from ooodev.formatters.formatter_table import FormatterTable, FormatTableItem
from ooodev.utils.file_io import FileIO
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.utils.type_var import PathOrStr, Row, Column

from ooo.dyn.sheet.cell_flags import CellFlags


class ExtractNums:
    def __init__(self, fnm: PathOrStr) -> None:
        _ = FileIO.is_exist_file(fnm, True)
        self._fnm = FileIO.get_absolute_path(fnm)

    def main(self) -> None:
        loader = Lo.load_office(Lo.ConnectSocket())

        try:
            doc = Calc.open_doc(fnm=self._fnm, loader=loader)

            GUI.set_visible(is_visible=True, odoc=doc)

            sheet = Calc.get_sheet(doc=doc, index=0)

            # basic data extraction
            # this code assumes the input file is "small totals.ods"
            print()
            print(f'A1 string: {Calc.get_val(sheet=sheet, cell_name="A1")}')  # string

            cell_name = "A2"
            cell = Calc.get_cell(sheet=sheet, cell_name=cell_name)
            print(f"{cell_name} type: {Calc.get_type_string(cell)}")
            print(f"{cell_name} value: {Calc.get_num(sheet=sheet, cell_name=cell_name)}")  # float

            cell_name = "E2"
            cell = Calc.get_cell(sheet=sheet, cell_name=cell_name)
            print(f"{cell_name} type: {Calc.get_type_string(cell)}")
            print(f"{cell_name} value: {Calc.get_val(sheet=sheet, cell_name=cell_name)}")  # formula string

            data = Calc.get_array(sheet=sheet, range_name="A1:E10")
            # apply formatting entire table except for first and last rows.
            # format as float with two decimal places.
            fl = FormatterTable(format=(".2f", ">9"), idxs=(0, 9))

            # add a custom row item formatter for first and last row only and pad items 9 spaces.
            fl.row_formats.append(FormatTableItem(format=">9", idxs_inc=(0, 9)))

            # add a custom columm formatter that formats the first column as integer values and move center in the column
            fl.col_formats.append(FormatTableItem(format=(".0f", "^9"), idxs_inc=(0,), row_idxs_exc=(0, 9)))

            # add a custom column formatter that formats the last columon as percent
            fl.col_formats.append(FormatTableItem(format=(".0%", ">9"), idxs_inc=(4,), row_idxs_exc=(0, 9)))
            Calc.print_array(data, fl)

            ids = Calc.get_float_array(sheet=sheet, range_name="A2:A7")
            fl = FormatterTable(format=(".1f", ">9"))
            Calc.print_array(ids, fl)

            projs = Calc.convert_to_floats(cast(Column, Calc.get_col(sheet=sheet, range_name="B2:B7")))
            print("Project scores")
            for proj in projs:
                print(f"  {proj:.2f}")

            stud = Calc.convert_to_floats(cast(Row, Calc.get_row(sheet=sheet, range_name="A4:E4")))
            print()
            print("Student scores")
            for v in stud:
                print(f"  {v:.2f}")

            # create a cell range that spans the used area of the sheet
            used_cell_rng = Calc.find_used_range(sheet)
            print()
            print(f"The used area is: {Calc.get_range_str(cell_range=used_cell_rng)}")

            # find cell ranges that cover all the specified data types
            cr_qry = Lo.qi(XCellRangesQuery, used_cell_rng)
            cell_ranges = cr_qry.queryContentCells(CellFlags.VALUE)
            # (CellFlags.VALUE | CellFlags.FORMULA)
            # (CellFlags.STRING)

            # process each of the cell ranges
            # -- extract each range as a 2D array of floats
            if cell_ranges is None:
                print("No cell ranges found")
            else:
                print(f"Found cell ranges: {cell_ranges.getRangeAddressesAsString()}")
                print()
                addrs = cell_ranges.getRangeAddresses()
                print(f'Cell Ranges: ({len(addrs)}):')
                fl = FormatterTable(format=(".2f", "<7"))
                # format the first col as integers
                fl.col_formats.append(FormatTableItem(format=(".0f", "<7"), idxs_inc=(0,)))
                for addr in addrs:
                    Calc.print_address(addr)
                    vals = Calc.get_float_array(sheet=sheet, range_name=Calc.get_range_str(addr))
                    print("WITH FORMATTING")
                    Calc.print_array(vals, fl)
                    # print("WITHOUT FORMATTING")
                    # Calc.print_array(vals)

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
