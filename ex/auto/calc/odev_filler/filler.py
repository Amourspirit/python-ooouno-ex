from __future__ import annotations

import uno

from ooodev.dialog.msgbox import (
    MessageBoxType,
    MessageBoxButtonsEnum,
    MessageBoxResultsEnum,
)
from ooodev.calc import Calc
from ooodev.calc import CalcDoc, CalcSheet
from ooodev.loader import Lo
from ooodev.utils.file_io import FileIO
from ooodev.utils.type_var import PathOrStr

from ooo.dyn.sheet.fill_date_mode import FillDateMode
from ooo.dyn.sheet.fill_direction import FillDirection
from ooo.dyn.sheet.fill_mode import FillMode


class Filler:
    def __init__(self, out_fnm: PathOrStr, **kwargs) -> None:
        if out_fnm:
            out_file = FileIO.get_absolute_path(out_fnm)
            _ = FileIO.make_directory(out_file)
            self._out_fnm = out_file
        else:
            self._out_fnm = ""

    def main(self) -> None:
        loader = Lo.load_office(Lo.ConnectSocket())

        try:
            doc = CalcDoc.create_doc(loader, visible=True)

            sheet = doc.sheets[0]

            self._fill_series(sheet)

            if self._out_fnm:
                doc.save_doc(fnm=self._out_fnm)

            msg_result = doc.msgbox(
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

    def _fill_series(self, sheet: CalcSheet) -> None:
        # set first two values of three rows
        sheet["B7"].value = 2
        sheet["A7"].value = 1  # 1. ascending

        sheet["A8"].set_date(day=28, month=2, year=2015)  # 2. dates, descending
        sheet["B8"].set_date(day=28, month=1, year=2015)

        sheet["A9"].value = 6  # 3. descending
        sheet["B9"].value = 4

        # Autofill, using first 2 cells to right to determine progressions
        series = sheet.get_range(range_name="A7:G9").get_cell_series()
        series.fillAuto(FillDirection.TO_RIGHT, 2)

        # ----------------------------------------
        sheet["A2"].value = 1
        sheet["A3"].value = 4

        # Fill 2 rows; 2nd row is not filled completely since
        # the end value is reached
        series = sheet.get_range(range_name="A2:E3").get_cell_series()
        series.fillSeries(FillDirection.TO_RIGHT, FillMode.LINEAR, Calc.NO_DATE, 2, 9)
        #   ignore date mode; step == 2; end at 9

        # ----------------------------------------
        sheet["A4"].set_date(day=20, month=11, year=2015)  # day, month, year
        # fill by adding one month to date; day is unchanged
        series = sheet.get_range(range_name="A4:E4").get_cell_series()
        series.fillSeries(
            FillDirection.TO_RIGHT,
            FillMode.DATE,
            FillDateMode.FILL_DATE_MONTH,
            1,
            Calc.MAX_VALUE,
        )
        # series.fillAuto(FillDirection.TO_RIGHT, 1)  # increments day not month

        # ----------------------------------------
        sheet["E5"].value = "Text 10"  # start in the middle of a row

        # Fill from right to left with text+value in steps of +10
        series = sheet.get_range(range_name="A5:E5").get_cell_series()
        series.fillSeries(
            FillDirection.TO_LEFT, FillMode.LINEAR, Calc.NO_DATE, 10, Calc.MAX_VALUE
        )

        # ----------------------------------------
        sheet["A6"].value = "Jan"

        # Fill with values generated automatically from first entry
        series = sheet.get_range(range_name="A6:E6").get_cell_series()
        series.fillSeries(
            FillDirection.TO_RIGHT, FillMode.AUTO, Calc.NO_DATE, 1, Calc.MAX_VALUE
        )
        # series.fillAuto(FillDirection.TO_RIGHT, 1)  # does the same

        # ----------------------------------------
        sheet["G6"].value = 10

        # Fill from  bottom to top with a geometric series (*2)
        series = sheet.get_range(range_name="G2:G6").get_cell_series()
        series.fillSeries(
            FillDirection.TO_TOP, FillMode.GROWTH, Calc.NO_DATE, 2, Calc.MAX_VALUE
        )
