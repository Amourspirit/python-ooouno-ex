from __future__ import annotations
from typing import Iterable

import uno
from com.sun.star.sheet import XSpreadsheet
from com.sun.star.table import XCellRange
from com.sun.star.util import XReplaceable
from com.sun.star.util import XSearchable

from ooodev.dialog.msgbox import MsgBox, MessageBoxType, MessageBoxButtonsEnum, MessageBoxResultsEnum
from ooodev.office.calc import Calc
from ooodev.utils.color import StandardColor
from ooodev.utils.file_io import FileIO
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.utils.table_helper import TableHelper
from ooodev.utils.type_var import PathOrStr
from ooodev.format import Styler
from ooodev.format.calc.direct.cell.background import Color as BgColor
from ooodev.format.calc.direct.cell.borders import Borders, BorderLineKind, Side, LineSize
from ooodev.format.calc.direct.cell.font import Font


class ReplaceAll:
    ANIMALS = (
        "ass",
        "cat",
        "cow",
        "cub",
        "doe",
        "dog",
        "elk",
        "ewe",
        "fox",
        "gnu",
        "hog",
        "kid",
        "kit",
        "man",
        "orc",
        "pig",
        "pup",
        "ram",
        "rat",
        "roe",
        "sow",
        "yak",
    )
    TOTAL_ROWS = 15
    TOTAL_COLS = 6

    def __init__(
        self,
        srch_strs: Iterable[str],
        repl_str: str | None = None,
        out_fnm: PathOrStr = "",
        is_search_all: bool = False,
    ) -> None:
        """
        Constructor

        Args:
            srch_strs (Iterable[str]): One or more search terms
            repl_str (str | None, optional): Replace String.
            out_fnm (PathOrStr, optional): File Save path.
            is_search_all (bool, optional): Determins if search is done in a iter manor or find all manor. Defaults to False (iter manor).
        """
        self._srch_strs = [s for s in srch_strs]
        self._repl_str = repl_str
        self._is_search_all = is_search_all
        if out_fnm:
            outf = FileIO.get_absolute_path(out_fnm)
            _ = FileIO.make_directory(outf)
            self._out_fnm = outf
        else:
            self._out_fnm = ""

    def main(self) -> None:
        loader = Lo.load_office(Lo.ConnectSocket())

        try:
            doc = Calc.create_doc(loader)

            GUI.set_visible(visible=True, doc=doc)
            Lo.delay(300)
            GUI.zoom(view=GUI.ZoomEnum.ZOOM_150_PERCENT)

            sheet = Calc.get_sheet(doc=doc, index=0)

            def cb(row: int, col: int, prev) -> str:
                # call back function for make_2d_array, sets the value for the cell
                # return animals repeating until all cells are filled
                v = (row * ReplaceAll.TOTAL_COLS) + col

                a_len = len(ReplaceAll.ANIMALS)
                if v > a_len - 1:
                    i = v % a_len
                else:
                    i = v
                return ReplaceAll.ANIMALS[i]

            tbl = TableHelper.make_2d_array(num_rows=ReplaceAll.TOTAL_ROWS, num_cols=ReplaceAll.TOTAL_COLS, val=cb)
            
            # create styles that can be applied to the cells via Calc.set_array_range().
            inner_side = Side()
            outer_side = Side(width=LineSize.THICK)
            bdr = Borders(border_side=outer_side, vertical=inner_side, horizontal=inner_side)
            bg_color = BgColor(StandardColor.BLUE)
            ft = Font(color=StandardColor.WHITE)

            Calc.set_array_range(sheet=sheet, range_name="A1:F15", values=tbl, styles=[bdr, bg_color, ft])

            # A1:F15
            cell_rng = Calc.get_cell_range(sheet=sheet, col_start=0, row_start=0, col_end=5, row_end=15)

            for s in self._srch_strs:
                if self._is_search_all:
                    self._search_all(sheet=sheet, cell_rng=cell_rng, srch_str=s)
                else:
                    self._search_iter(sheet=sheet, cell_rng=cell_rng, srch_str=s)

            if self._repl_str is not None:
                for s in self._srch_strs:
                    self._replace_all(cell_rng=cell_rng, srch_str=s, repl_str=self._repl_str)

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

    def _highlight(self, cr: XCellRange) -> None:
        # highlight by make cell bold, with text color of Light purple and a background color of light blue.
        ft = Font(b=True, color=StandardColor.PURPLE_LIGHT1)
        bg_color = BgColor(StandardColor.DEFAULT_BLUE)
        borders = Borders(border_side=Side(line=BorderLineKind.SOLID, color=StandardColor.RED_DARK3))
        Styler.apply(cr, ft, bg_color, borders)

    def _search_iter(self, sheet: XSpreadsheet, cell_rng: XCellRange, srch_str: str) -> None:
        print(f'Searching (iterating) for all occurrences of "{srch_str}"')
        try:
            srch = Lo.qi(XSearchable, cell_rng, True)
            sd = srch.createSearchDescriptor()

            sd.setSearchString(srch_str)
            # only complete words will be found
            sd.setPropertyValue("SearchWords", True)
            # sd.setPropertyValue("SearchRegularExpression", True)

            o_first = srch.findFirst(sd)
            # Info.show_services("Find First", o_first)

            cr = Lo.qi(XCellRange, o_first)
            if cr is None:
                print(f'  No match found for "{srch_str}"')
                return
            count = 0
            while cr is not None:
                self._highlight(cr)
                print(f"  Match {count + 1} : {Calc.get_range_str(cr)}")
                cr = Lo.qi(XCellRange, srch.findNext(cr, sd))
                count += 1

        except Exception as e:
            print(e)

    def _search_all(self, sheet: XSpreadsheet, cell_rng: XCellRange, srch_str: str) -> None:
        print(f'Searching (find all) for all occurrences of "{srch_str}"')
        try:
            srch = Lo.qi(XSearchable, cell_rng, True)
            sd = srch.createSearchDescriptor()

            sd.setSearchString(srch_str)
            sd.setPropertyValue("SearchWords", True)

            match_crs = Calc.find_all(srch=srch, sd=sd)
            if not match_crs:
                print(f'  No match found for "{srch_str}"')
                return
            for i, cr in enumerate(match_crs):
                self._highlight(cr)
                print(f"  Index {i} : {Calc.get_range_str(cr)}")

        except Exception as e:
            print(e)

    def _replace_all(self, cell_rng: XCellRange, srch_str: str, repl_str: str) -> None:
        print(f'Replacing "{srch_str}" with "{repl_str}"')
        Lo.delay(2000)  # wait a bit before search & replace
        try:
            repl = Lo.qi(XReplaceable, cell_rng, True)
            rd = repl.createReplaceDescriptor()

            rd.setSearchString(srch_str)
            rd.setReplaceString(repl_str)
            rd.setPropertyValue("SearchWords", True)
            # rd.setPropertyValue("SearchRegularExpression", True)

            count = repl.replaceAll(rd)
            print(f"Search text replaced {count} times")
            print()

        except Exception as e:
            print(e)
