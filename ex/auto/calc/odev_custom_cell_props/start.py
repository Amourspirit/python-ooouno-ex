#!/usr/bin/env python
from __future__ import annotations
from typing import Any
from pathlib import Path
import shutil
import logging
import uno
from com.sun.star.document import MacroExecMode

from ooodev.io.log import logging as logger
from ooodev.loader.inst.options import Options
from ooodev.calc import CalcDoc, ZoomKind
from ooodev.calc.cell.custom_prop_clean import CustomPropClean
from ooodev.loader import Lo
from ooodev.dialog.msgbox import (
    MessageBoxType,
    MessageBoxButtonsEnum,
    MessageBoxResultsEnum,
)

import macro_code


def clean_doc() -> None:
    doc = CalcDoc.from_current_doc()
    for sheets in doc.sheets:
        cleaner = CustomPropClean(sheets)
        cleaner.clean()


def doc_event_occurred(src: Any, args: Any, control_src: Any) -> None:
    # print("Event occurred")
    name = args.event_data.EventName
    if name in ("OnCopyTo", "OnSaveAs", "OnSave"):
        clean_doc()
        logger.debug("Cleaned")


def main() -> int:
    _ = Lo.load_office(Lo.ConnectSocket(), opt=Options(log_level=logging.DEBUG))
    try:
        fnm = Path(__file__).parent / "data" / "custom_props.ods"  # "custom_props.ods"
        tmp_dir = Lo.tmp_dir / "custom_cell_props"
        tmp_dir.mkdir(parents=True, exist_ok=True)
        dest_file = tmp_dir / fnm.name
        shutil.copy(fnm, dest_file)

        doc: CalcDoc = CalcDoc.open_doc(
            fnm=dest_file,
            visible=True,
            MacroExecutionMode=MacroExecMode.ALWAYS_EXECUTE_NO_WARN,
        )
        Lo.global_event_broadcaster.add_event_document_event_occurred(
            doc_event_occurred
        )
        # wait for document to be ready
        Lo.delay(300)
        doc.zoom(ZoomKind.ZOOM_100_PERCENT)
        macro_code.register_custom_prop_interceptor()

        # a breakpoint can be set below to inspect the menu.

        msg_result = doc.msgbox(
            "Do you wish to close document?",
            "All done",
            boxtype=MessageBoxType.QUERYBOX,
            buttons=MessageBoxButtonsEnum.BUTTONS_YES_NO,
        )
        if msg_result == MessageBoxResultsEnum.YES:
            doc.save_doc(fnm=dest_file)
            logger.debug(f"Saved document to: {dest_file}")
            doc.close()
            Lo.close_office()
        else:
            # un-register or expect a crash
            macro_code.unregister_custom_prop_interceptor()
            logger.debug("Keeping document open")
    except Exception:
        Lo.close_office()
        raise
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
