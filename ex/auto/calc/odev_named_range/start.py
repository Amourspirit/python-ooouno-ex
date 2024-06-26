#!/usr/bin/env python
from __future__ import annotations
from typing import cast, TYPE_CHECKING
from pathlib import Path
import shutil
import logging
import uno
from com.sun.star.document import MacroExecMode
from com.sun.star.table import CellAddress
from ooo.dyn.sheet.named_range_flag import NamedRangeFlagEnum

from ooodev.io.log import logging as logger
from ooodev.loader.inst.options import Options
from ooodev.calc import CalcDoc, CalcSheet, ZoomKind
from ooodev.calc.cell.custom_prop_clean import CustomPropClean
from ooodev.loader import Lo
from ooodev.dialog.msgbox import (
    MessageBoxType,
    MessageBoxButtonsEnum,
    MessageBoxResultsEnum,
)

if TYPE_CHECKING:
    from com.sun.star.sheet import SheetCellRange  # service


def get_named_range(sheet: CalcSheet, name: str) -> SheetCellRange | None:
    nc = None
    if sheet.named_ranges.has_by_name(name):
        nc = sheet.named_ranges.get_by_name(name)
    elif sheet.calc_doc.named_ranges.has_by_name(name):
        nc = sheet.calc_doc.named_ranges.get_by_name(name)
    if nc is None:
        return None
    return cast("SheetCellRange", nc.get_referred_cells())


def main() -> int:
    _ = Lo.load_office(Lo.ConnectSocket(), opt=Options(log_level=logging.DEBUG))
    try:
        fnm = Path(__file__).parent / "data" / "produceSales.ods"

        doc: CalcDoc = CalcDoc.open_doc(fnm=fnm, visible=True)

        # wait for document to be ready
        Lo.delay(300)
        doc.zoom(ZoomKind.ZOOM_100_PERCENT)

        sheet = doc.sheets[0]
        ca = sheet.get_cell_address(cell_name="A2")
        sheet_named_rngs = sheet.named_ranges
        sheet_named_rngs.add_new_by_name(
            name="my_sheet_range", content="$Sheet.$A$2:$D$5", position=ca
        )

        ca = sheet.get_cell_address(cell_name="A7")
        doc_named_rngs = doc.named_ranges
        doc_named_rngs.add_new_by_name(
            name="my_doc_range", content="$Sheet.$A$7:$D$11", position=ca
        )

        nr = get_named_range(sheet, "my_sheet_range")
        if nr is not None:
            rng_obj = doc.range_converter.get_range_obj(nr.getRangeAddress())
            logger.debug(f"Sheet Range: {rng_obj}")
            sheet.select_cells_range(rng_obj)

        msg_result = doc.msgbox(
            "Do you wish to remove Sheet Range?",
            "Remove Sheet Range",
            boxtype=MessageBoxType.QUERYBOX,
            buttons=MessageBoxButtonsEnum.BUTTONS_YES_NO,
        )

        if msg_result == MessageBoxResultsEnum.YES:
            sheet_named_rngs.remove_by_name("my_sheet_range")

        msg_result = doc.msgbox(
            "Do you wish to remove Document Range?",
            "Remove Document Range",
            boxtype=MessageBoxType.QUERYBOX,
            buttons=MessageBoxButtonsEnum.BUTTONS_YES_NO,
        )

        if msg_result == MessageBoxResultsEnum.YES:
            doc_named_rngs.remove_by_name("my_doc_range")

        msg_result = doc.msgbox(
            "Do you wish to close document?",
            "All done",
            boxtype=MessageBoxType.QUERYBOX,
            buttons=MessageBoxButtonsEnum.BUTTONS_YES_NO,
        )
        if msg_result == MessageBoxResultsEnum.YES:
            doc.close()
            Lo.close_office()
        else:
            logger.debug("Keeping document open")
    except Exception:
        Lo.close_office()
        raise
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
