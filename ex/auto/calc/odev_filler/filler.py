from __future__ import annotations

import uno
from com.sun.star.sheet import XSpreadsheet

from ooodev.dialog.msgbox import MsgBox, MessageBoxType, MessageBoxButtonsEnum, MessageBoxResultsEnum
from ooodev.office.calc import Calc
from ooodev.utils.file_io import FileIO
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.utils.type_var import PathOrStr

from ooo.dyn.sheet.fill_date_mode import FillDateMode
from ooo.dyn.sheet.fill_direction import FillDirection
from ooo.dyn.sheet.fill_mode import FillMode


class Filler:
    def __init__(self, out_fnm: PathOrStr, **kwargs) -> None:
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

            GUI.set_visible(is_visible=True, odoc=doc)

            sheet = Calc.get_sheet(doc=doc, index=0)

            self._fill_series(sheet)

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

    def _fill_series(self, sheet: XSpreadsheet) -> None:
        # set first two values of three rows
        Calc.set_val(sheet=sheet, cell_name="B7", value=2)
        Calc.set_val(sheet=sheet, cell_name="A7", value=1)  # 1. ascending

        Calc.set_date(sheet=sheet, cell_name="A8", day=28, month=2, year=2015)  # 2. dates, descending
        Calc.set_date(sheet=sheet, cell_name="B8", day=28, month=1, year=2015)

        Calc.set_val(sheet=sheet, cell_name="A9", value=6)  # 3. descending
        Calc.set_val(sheet=sheet, cell_name="B9", value=4)

        # Autofill, using first 2 cells to right to determine progressions
        series = Calc.get_cell_series(sheet=sheet, range_name="A7:G9")
        series.fillAuto(FillDirection.TO_RIGHT, 2)

        # ----------------------------------------
        Calc.set_val(sheet=sheet, cell_name="A2", value=1)
        Calc.set_val(sheet=sheet, cell_name="A3", value=4)

        # Fill 2 rows; 2nd row is not filled completely since
        # the end value is reached
        series = Calc.get_cell_series(sheet=sheet, range_name="A2:E3")
        series.fillSeries(FillDirection.TO_RIGHT, FillMode.LINEAR, Calc.NO_DATE, 2, 9)
        #   ignore date mode; step == 2; end at 9

        # ----------------------------------------
        Calc.set_date(sheet=sheet, cell_name="A4", day=20, month=11, year=2015)  # day, month, year
        # fill by adding one month to date; day is unchanged
        series = Calc.get_cell_series(sheet=sheet, range_name="A4:E4")
        series.fillSeries(FillDirection.TO_RIGHT, FillMode.DATE, FillDateMode.FILL_DATE_MONTH, 1, Calc.MAX_VALUE)
        # series.fillAuto(FillDirection.TO_RIGHT, 1)  # increments day not month

        # ----------------------------------------
        Calc.set_val(sheet=sheet, cell_name="E5", value="Text 10")  # start in the middle of a row

        # Fill from right to left with text+value in steps of +10
        series = Calc.get_cell_series(sheet=sheet, range_name="A5:E5")
        series.fillSeries(FillDirection.TO_LEFT, FillMode.LINEAR, Calc.NO_DATE, 10, Calc.MAX_VALUE)

        # ----------------------------------------
        Calc.set_val(sheet=sheet, cell_name="A6", value="Jan")

        # Fill with values generated automatically from first entry
        series = Calc.get_cell_series(sheet=sheet, range_name="A6:E6")
        series.fillSeries(FillDirection.TO_RIGHT, FillMode.AUTO, Calc.NO_DATE, 1, Calc.MAX_VALUE)
        # series.fillAuto(FillDirection.TO_RIGHT, 1)  # does the same

        # ----------------------------------------
        Calc.set_val(sheet=sheet, cell_name="G6", value=10)

        # Fill from  bottom to top with a geometric series (*2)
        series = Calc.get_cell_series(sheet=sheet, range_name="G2:G6")
        series.fillSeries(FillDirection.TO_TOP, FillMode.GROWTH, Calc.NO_DATE, 2, Calc.MAX_VALUE)
