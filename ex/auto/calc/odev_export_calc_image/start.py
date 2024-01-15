from __future__ import annotations
from typing import cast

import uno
from com.sun.star.sheet import XCellRangesQuery
from ooo.dyn.sheet.cell_flags import CellFlags

from ooodev.dialog.msgbox import (
    MsgBox,
    MessageBoxType,
    MessageBoxButtonsEnum,
    MessageBoxResultsEnum,
)
from ooodev.calc import Calc, CalcDoc, ZoomKind
from ooodev.formatters.formatter_table import FormatterTable, FormatTableItem
from ooodev.utils.file_io import FileIO
from ooodev.utils.lo import Lo
from ooodev.utils.type_var import PathOrStr, Row, Column
from pathlib import Path


def main() -> int:
    fnm = Path(__file__).parent / "data" / "data.ods"

    loader = Lo.load_office(Lo.ConnectSocket())

    try:
        doc = CalcDoc(Calc.open_doc(fnm=fnm, loader=loader))

        doc.set_visible()
        # delay before dispatching zoom
        Lo.delay(500)
        doc.zoom(ZoomKind.ZOOM_100_PERCENT)

        sheet = doc.get_active_sheet()

        # get a range of cells from the sheet
        rng = sheet.get_range(range_name="A1:N4")

        # create a path to store file
        tmp = Path.cwd() / "tmp"
        tmp.mkdir(exist_ok=True)
        file = tmp / "data.png"

        # export range as image, in this case a png.
        # jpg format is also supported, eg: data.jpg
        rng.export_as_image(file)

        # quick check to see if file exists
        assert file.exists()
        print(f"Image exported to {file}")

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
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
