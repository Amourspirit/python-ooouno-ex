from __future__ import annotations

import uno
from com.sun.star.sheet import XCellRangesQuery
from com.sun.star.sheet import XSpreadsheet
from com.sun.star.sheet import XSpreadsheetDocument
from com.sun.star.util import XMergeable

from ooodev.dialog.msgbox import MsgBox, MessageBoxType, MessageBoxButtonsEnum, MessageBoxResultsEnum
from ooodev.office.calc import Calc, GeneralFunction
from ooodev.utils.color import CommonColor
from ooodev.utils.file_io import FileIO
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.utils.props import Props
from ooodev.utils.type_var import PathOrStr
from ooodev.utils.view_state import ViewState

from ooo.dyn.awt.font_weight import FontWeight
from ooo.dyn.table.cell_content_type import CellContentType
from ooo.dyn.table.cell_hori_justify import CellHoriJustify
from ooo.dyn.table.cell_vert_justify import CellVertJustify


class GarlicSecrets:
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

            sheet = Calc.get_sheet(doc=doc, index=0)
            Calc.goto_cell(cell_name="A1", doc=doc)

            # freeze one row of view
            # Calc.freeze_rows(doc=doc, num_rows=1)

            # find total for the "Total" column
            total_range = Calc.get_col_range(sheet=sheet, idx=3)
            total = Calc.compute_function(fn=GeneralFunction.SUM, cell_range=total_range)
            print(f"Total before change: {total:.2f}")
            print()

            self._increase_garlic_cost(doc=doc, sheet=sheet)  # takes several seconds

            # recalculate total
            total = Calc.compute_function(fn=GeneralFunction.SUM, cell_range=total_range)
            print(f"Total after change: {total:.2f}")
            print()

            # add a label at the bottom of the data, and hide it
            empty_row_num = self._find_empty_row(sheet=sheet)
            self._add_garlic_label(doc=doc, sheet=sheet, empty_row_num=empty_row_num)
            Lo.delay(2_000)  # wait a bit before hiding last row
            row_range = Calc.get_row_range(sheet=sheet, idx=empty_row_num)
            Props.set(row_range, IsVisible=False)  # Property is in TableRow

            # split window into 2 view panes
            cell_name = Calc.get_cell_str(col=0, row=empty_row_num - 2)
            print(f"Splitting at: {cell_name}")
            # doesn't work with Calc.freeze()
            Calc.split_window(doc=doc, cell_name=cell_name)

            # access panes; make top pane show first row
            panes = Calc.get_view_panes(doc=doc)
            print(f"No. of panes: {len(panes)}")
            panes[0].setFirstVisibleRow(0)

            # display view properties
            ss_view = Calc.get_view(doc=doc)
            Props.show_obj_props("Spreadsheet view", ss_view)

            # show view data
            print(f"View data: {Calc.get_view_data(doc)}")

            # show sheet states
            states = Calc.get_view_states(doc=doc)
            for s in states:
                s.report()

            # make top pane the active one in the first sheet
            states[0].move_pane_focus(dir=ViewState.PaneEnum.MOVE_UP)
            Calc.set_view_states(doc=doc, states=states)
            # move selection to top cell
            Calc.goto_cell(cell_name="A1", doc=doc)

            # show revised sheet states
            states = Calc.get_view_states(doc=doc)
            for s in states:
                s.report()

            # add a new first row, and label that as at the bottom
            Calc.insert_row(sheet=sheet, idx=0)
            self._add_garlic_label(doc=doc, sheet=sheet, empty_row_num=0)

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

    def _increase_garlic_cost(self, doc: XSpreadsheetDocument, sheet: XSpreadsheet) -> int:
        """
        Iterate down the "Produce" column. If the text in the current cell is
        "Garlic" then change the corresponding "Cost Per Pound" cell by
        multiplying by 1.05, and changing its text to bold red.

        Return the "Produce" row index which is first empty.
        """

        row = 0
        prod_cell = Calc.get_cell(sheet=sheet, col=0, row=row)  # produce column
        # iterate down produce column until an empty cell is reached
        while prod_cell.getType() != CellContentType.EMPTY:
            if prod_cell.getFormula() == "Garlic":
                # show the cell in-screen
                Calc.goto_cell(doc=doc, cell_name=Calc.get_cell_str(col=0, row=row))
                # change cost/pound column
                cost_cell = Calc.get_cell(sheet=sheet, col=1, row=row)
                cost_cell.setValue(1.05 * cost_cell.getValue())
                Props.set(cost_cell, CharWeight=FontWeight.BOLD, CharColor=CommonColor.RED)
            row += 1
            prod_cell = Calc.get_cell(sheet=sheet, col=0, row=row)
        return row

    def _find_empty_row(self, sheet: XSpreadsheet) -> int:
        """
        Return the index of the first empty row by finding all the empty cell ranges in
        the first column, and return the smallest row index in those ranges.
        """

        # create a ranges query for the first column of the sheet
        cell_range = Calc.get_col_range(sheet=sheet, idx=0)
        Calc.print_address(cell_range=cell_range)
        cr_query = Lo.qi(XCellRangesQuery, cell_range)
        sc_ranges = cr_query.queryEmptyCells()
        addrs = sc_ranges.getRangeAddresses()
        Calc.print_addresses(*addrs)

        # find smallest row index
        row = -1
        if addrs is not None and len(addrs) > 0:
            row = addrs[0].StartRow
            for addr in addrs:
                if row < addr.StartRow:
                    row = addr.StartRow
            print(f"First empty row is at position: {row}")
        else:
            print("Could not find an empty row")
        return row

    def _add_garlic_label(self, doc: XSpreadsheetDocument, sheet: XSpreadsheet, empty_row_num: int) -> None:
        """
        Add a large text string ("Top Secret Garlic Changes") to the first cell
        in the empty row. Make the cell bigger by merging a few cells, and taller
        The text is black and bold in a red cell, and is centered.
        """

        Calc.goto_cell(cell_name=Calc.get_cell_str(col=0, row=empty_row_num), doc=doc)

        # Merge first few cells of the last row
        cell_range = Calc.get_cell_range(
            sheet=sheet, start_col=0, start_row=empty_row_num, end_col=3, end_row=empty_row_num
        )
        xmerge = Lo.qi(XMergeable, cell_range, True)
        xmerge.merge(True)

        # make the row taller
        Calc.set_row_height(sheet=sheet, height=18, idx=empty_row_num)
        cell = Calc.get_cell(sheet=sheet, col=0, row=empty_row_num)
        cell.setFormula("Top Secret Garlic Changes")
        Props.set(
            cell,
            CharWeight=FontWeight.BOLD,
            CharHeight=24,
            CellBackColor=CommonColor.RED,
            HoriJustify=CellHoriJustify.CENTER,
            VertJustify=CellVertJustify.CENTER,
        )
