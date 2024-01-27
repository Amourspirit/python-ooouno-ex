from __future__ import annotations
import time

import uno
from com.sun.star.sheet import XCellRangesQuery

from ooo.dyn.table.cell_content_type import CellContentType

from ooodev.dialog.msgbox import (
    MsgBox,
    MessageBoxType,
    MessageBoxButtonsEnum,
    MessageBoxResultsEnum,
)
from ooodev.calc import Calc, GeneralFunction
from ooodev.calc import CalcDoc, CalcSheet
from ooodev.format.calc.direct.cell.background import Color as BgColor
from ooodev.format.calc.direct.cell.font import Font
from ooodev.utils.color import CommonColor
from ooodev.utils.file_io import FileIO
from ooodev.utils.lo import Lo
from ooodev.utils.props import Props
from ooodev.utils.type_var import PathOrStr
from ooodev.utils.view_state import ViewState


class GarlicSecrets:
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
            doc = CalcDoc.open_doc(fnm=self._fnm, loader=loader)

            doc.set_visible()

            sheet = doc.sheets[0]
            sheet.goto_cell(cell_name="A1")

            # freeze one row of view
            # doc.freeze_rows(num_rows=1)

            # find total for the "Total" column
            total_range = sheet.get_col_range(idx=3)
            total = doc.compute_function(
                fn=GeneralFunction.SUM, cell_range=total_range.component
            )
            print(f"Total before change: {total:.2f}")
            print()

            # there are 3 ways to iterate over the data
            # the first two take about the same time and are longer.
            # self._increase_garlic_cost1(sheet)
            # self._increase_garlic_cost2(sheet)
            self._increase_garlic_cost3(sheet)  # About %30 faster

            # recalculate total
            total = doc.compute_function(fn=GeneralFunction.SUM, cell_range=total_range)
            print(f"Total after change: {total:.2f}")
            print()

            # add a label at the bottom of the data, and hide it
            empty_row_num = self._find_empty_row(sheet)
            self._add_garlic_label(sheet=sheet, empty_row_num=empty_row_num)
            Lo.delay(2_000)  # wait a bit before hiding last row
            row_range = sheet.get_row_range(idx=empty_row_num)
            row_range.is_visible = False

            # split window into 2 view panes
            cell_name = sheet[(0, empty_row_num - 2)].get_cell_str()
            print(f"Splitting at: {cell_name}")
            # doesn't work with Calc.freeze()
            doc.split_window(cell_name=cell_name)

            # access panes; make top pane show first row
            panes = doc.get_view_panes()
            if panes:
                print(f"No. of panes: {len(panes)}")
                panes[0].setFirstVisibleRow(0)

            # display view properties
            ss_view = doc.get_view()
            Props.show_obj_props("Spreadsheet view", ss_view.component)

            # show view data
            print(f"View data: {doc.get_view_data()}")

            # show sheet states
            states = doc.get_view_states()
            if states:
                for s in states:
                    s.report()

                # make top pane the active one in the first sheet
                states[0].move_pane_focus(dir=ViewState.PaneEnum.MOVE_UP)
                doc.set_view_states(states=states)
            # move selection to top cell
            sheet["A1"].goto()

            # show revised sheet states
            states = doc.get_view_states()
            if states:
                for s in states:
                    s.report()

            # add a new first row, and label that as at the bottom
            sheet.insert_row(idx=0)
            self._add_garlic_label(sheet=sheet, empty_row_num=0)

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

    def _increase_garlic_cost1(self, sheet: CalcSheet) -> int:
        """
        Iterate down the "Produce" column. If the text in the current cell is
        "Garlic" then change the corresponding "Cost Per Pound" cell by
        multiplying by 1.05, and changing its text to bold red.

        Return the "Produce" row index which is first empty.

        Both _increase_garlic_cost1() and _increase_garlic_cost2() take around the same time
        with _increase_garlic_cost1() being slightly faster about %1 difference.
        The _increase_garlic_cost3() method is about 30% faster.
        """
        start_time = time.time()
        row = 0
        cell = sheet[(0, row)]  # produce column
        sht_comp = sheet.component
        prod_cell = cell.component
        red_font = Font(b=True, color=CommonColor.RED)

        # iterate down produce column until an empty cell is reached
        while prod_cell.getType() != CellContentType.EMPTY:
            if prod_cell.getFormula() == "Garlic":
                # show the cell in-screen
                cell = sheet[(0, row)]
                cell.goto()
                # change cost/pound column
                cost_cell = sheet[(1, row)]
                # make the change more visible by making the text bold and red
                cost_cell.set_val(1.05 * cost_cell.get_num(), [red_font])
            row += 1

            # this next line will work but is slower
            # prod_cell = sheet.get_cell(col=0, row=row).component
            # sheet.get_cell() will create a CallCell object, which is slower then getting
            # XCell directly from Calc.get_cell().
            # When iterating over a large number of cells, it is better to use Calc.get_cell().
            prod_cell = Calc.get_cell(sheet=sht_comp, col=0, row=row)
        end_time = time.time()
        print(f"Time to iterate over array: {end_time - start_time:.2f} seconds")
        return row

    def _increase_garlic_cost2(self, sheet: CalcSheet) -> int:
        """
        Iterate down the "Produce" column. If the text in the current cell is
        "Garlic" then change the corresponding "Cost Per Pound" cell by
        multiplying by 1.05, and changing its text to bold red.

        Return the "Produce" row index which is first empty.

        Both _increase_garlic_cost1() and _increase_garlic_cost2() take around the same time
        with _increase_garlic_cost1() being slightly faster about %1 difference.
        The _increase_garlic_cost3() method is about 30% faster.
        """
        start_time = time.time()
        row = 0
        cell = sheet.get_cell(col=0, row=row)  # produce column
        sht_comp = sheet.component
        prod_cell = cell.component
        red_font = Font(b=True, color=CommonColor.RED)

        used_rng = sheet.find_used_range_obj()
        rng_col1 = 1 - used_rng.get_col(0)  # omit the first row
        for row, cells in enumerate(rng_col1.get_cells()):
            row += 1
            for cell_obj in cells:
                prod_cell = Calc.get_cell(sheet=sht_comp, cell_obj=cell_obj)
                if prod_cell.getFormula() == "Garlic":
                    cell = sheet[(0, row)]
                    cell.goto()
                    # change cost/pound column
                    cost_cell = sheet[(1, row)]
                    # make the change more visible by making the text bold and red
                    cost_cell.set_val(1.05 * cost_cell.get_num(), [red_font])
        end_time = time.time()
        print(f"Time to iterate over array: {end_time - start_time:.2f} seconds")
        return row

    def _increase_garlic_cost3(self, sheet: CalcSheet) -> int:
        """
        Iterate down the "Produce" column. If the text in the current cell is
        "Garlic" then change the corresponding "Cost Per Pound" cell by
        multiplying by 1.05, and changing its text to bold red.

        Return the "Produce" row index which is first empty.

        This method is the fastest. It get the data as an array, and then iterates over the array.

        Both _increase_garlic_cost1() and _increase_garlic_cost2() take around the same time
        with _increase_garlic_cost1() being slightly faster about %1 difference.

        This method is about 30% faster than _increase_garlic_cost1() and _increase_garlic_cost2().
        This method is faster because it gets the data as an array, and then iterates over the array.
        The result is that the number of calls to the UNO interface is reduced greatly.
        """
        start_time = time.time()
        row_count = 0
        cell = sheet.get_cell(col=0, row=row_count)  # produce column
        red_font = Font(b=True, color=CommonColor.RED)

        used_rng = sheet.find_used_range_obj()
        rng_col1 = 1 - used_rng.get_col(0)  # omit the first row (header)
        data = sheet.get_array(range_obj=rng_col1)
        for row in data:
            row_count += 1
            if row[0] == "Garlic":
                cell = sheet[(0, row_count)]
                cell.goto()
                # change cost/pound column
                cost_cell = sheet[(1, row_count)]
                # make the change more visible by making the text bold and red
                cost_cell.set_val(1.05 * cost_cell.get_num(), [red_font])
        end_time = time.time()
        print(f"Time to iterate over array: {end_time - start_time:.2f} seconds")

        return row_count

    def _find_empty_row(self, sheet: CalcSheet) -> int:
        """
        Return the index of the first empty row by finding all the empty cell ranges in
        the first column, and return the smallest row index in those ranges.
        """

        # create a ranges query for the first column of the sheet
        cell_range = sheet.get_col_range(idx=0)
        Calc.print_address(cell_range=cell_range.component)
        cr_query = cell_range.qi(XCellRangesQuery, True)
        sc_ranges = cr_query.queryEmptyCells()
        addresses = sc_ranges.getRangeAddresses()
        Calc.print_addresses(*addresses)

        # find smallest row index
        row = -1
        if addresses is not None and len(addresses) > 0:
            row = addresses[0].StartRow
            for addr in addresses:
                if row < addr.StartRow:
                    row = addr.StartRow
            print(f"First empty row is at position: {row}")
        else:
            print("Could not find an empty row")
        return row

    def _add_garlic_label(self, sheet: CalcSheet, empty_row_num: int) -> None:
        """
        Add a large text string ("Top Secret Garlic Changes") to the first cell
        in the empty row. Make the cell bigger by merging a few cells, and taller
        The text is black and bold in a red cell, and is centered.
        """

        sheet.goto_cell(cell_obj=sheet[(0, empty_row_num)].cell_obj)

        # Merge first few cells of the last row
        rng_obj = Calc.get_range_obj(
            col_start=0, row_start=empty_row_num, col_end=3, row_end=empty_row_num
        )

        # merge and center range
        sheet.merge_cells(range_obj=rng_obj, center=True)

        # make the row taller
        sheet.set_row_height(height=18, idx=empty_row_num)
        # get the cell from the range cell start
        font_red = Font(b=True, size=24, color=CommonColor.BLACK)
        bg_color = BgColor(CommonColor.RED)

        cell = sheet[rng_obj.cell_start]
        cell.set_val(value="Top Secret Garlic Changes", styles=[font_red, bg_color])
