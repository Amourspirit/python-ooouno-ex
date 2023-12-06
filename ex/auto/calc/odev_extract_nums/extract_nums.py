from __future__ import annotations
from typing import cast

import uno
from com.sun.star.sheet import XCellRangesQuery

from ooodev.dialog.msgbox import (
    MsgBox,
    MessageBoxType,
    MessageBoxButtonsEnum,
    MessageBoxResultsEnum,
)
from ooodev.calc import Calc
from ooodev.calc import CalcDoc
from ooodev.formatters.formatter_table import FormatterTable, FormatTableItem
from ooodev.utils.file_io import FileIO
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
            doc = CalcDoc(Calc.open_doc(fnm=self._fnm, loader=loader))

            doc.set_visible()

            sheet = doc.get_active_sheet()

            # basic data extraction
            # this code assumes the input file is "small totals.ods"
            print()
            print(f'A1 string: {sheet.get_val(cell_name="A1")}')  # string

            cell_name = "A2"
            cell = sheet.get_cell(cell_name=cell_name)
            print(f"{cell_name} type: {cell.get_type_string()}")
            print(f"{cell_name} value: {sheet.get_num(cell_name=cell_name)}")  # float

            cell_name = "E2"
            cell = sheet.get_cell(cell_name=cell_name)
            print(f"{cell_name} type: {cell.get_type_string()}")
            print(f"{cell_name} value: {cell.get_val()}")  # formula string

            rng = sheet.get_range(range_name="A1:E10")
            data = rng.get_array()
            # apply formatting entire table except for first and last rows.
            start_idx = Calc.get_row_used_first_index(sheet.component)
            end_idx = Calc.get_row_used_last_index(sheet.component)
            # format as float with two decimal places.
            fl = FormatterTable(format=(".2f", ">9"), idxs=(start_idx, end_idx))

            # add a custom row item formatter for first and last row only and pad items 9 spaces.
            fl.row_formats.append(
                FormatTableItem(format=">9", idxs_inc=(start_idx, end_idx))
            )

            # add a custom column formatter that formats the first column as integer values and move center in the column
            fl.col_formats.append(
                FormatTableItem(
                    format=(".0f", "^9"),
                    idxs_inc=(start_idx,),
                    row_idxs_exc=(start_idx, end_idx),
                )
            )

            # add a custom column formatter that formats the last column as percent
            fl.col_formats.append(
                FormatTableItem(
                    format=(".0%", ">9"),
                    idxs_inc=(4,),
                    row_idxs_exc=(start_idx, end_idx),
                )
            )
            Calc.print_array(data, fl)

            ids = sheet.get_float_array(range_name="A2:A7")
            fl = FormatterTable(format=(".1f", ">9"))
            Calc.print_array(ids, fl)

            projects = Calc.convert_to_floats(
                cast(Column, sheet.get_col(range_name="B2:B7"))
            )
            print("Project scores")
            for proj in projects:
                print(f"  {proj:.2f}")

            stud = Calc.convert_to_floats(cast(Row, sheet.get_row(range_name="A4:E4")))
            print()
            print("Student scores")
            for v in stud:
                print(f"  {v:.2f}")

            # create a cell range that spans the used area of the sheet
            used_cell_rng = sheet.find_used_range()
            print()
            print(f"The used area is: {used_cell_rng.get_range_str()}")

            # find cell ranges that cover all the specified data types
            cr_qry = used_cell_rng.qi(XCellRangesQuery, True)
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
                addresses = cell_ranges.getRangeAddresses()
                print(f"Cell Ranges: ({len(addresses)}):")
                fl = FormatterTable(format=(".2f", "<7"))
                # format the first col as integers
                fl.col_formats.append(
                    FormatTableItem(format=(".0f", "<7"), idxs_inc=(start_idx,))
                )
                for addr in addresses:
                    Calc.print_address(addr)
                    vals = sheet.get_float_array(range_name=Calc.get_range_str(addr))
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
                doc.close_doc()
                Lo.close_office()
            else:
                print("Keeping document open")

        except Exception:
            Lo.close_office()
            raise
