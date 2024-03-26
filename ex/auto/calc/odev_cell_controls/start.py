#!/usr/bin/env python
from __future__ import annotations
from pathlib import Path
import uno
from com.sun.star.document import MacroExecMode

from ooodev.loader import Lo
from ooodev.calc import CalcDoc, ZoomKind
from sheet_controls import SheetControls


def main() -> int:

    fnm = Path(__file__).parent / "data" / "odev_cell_controls.ods"
    _ = Lo.load_office(Lo.ConnectSocket())
    try:
        doc = CalcDoc.open_doc(
            fnm=fnm,
            visible=True,
            MacroExecutionMode=MacroExecMode.ALWAYS_EXECUTE_NO_WARN,
        )
        # wait for document to be ready
        Lo.delay(300)
        doc.zoom(ZoomKind.ZOOM_100_PERCENT)
        _ = SheetControls(doc)
        doc.msgbox("Sheet Controls added", "Sheet Controls", boxtype=1)
        Lo.delay(2_000)
        print("Keeping document open")
    except Exception:
        Lo.close_office()
        raise
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
