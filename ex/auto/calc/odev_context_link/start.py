#!/usr/bin/env python
from __future__ import annotations

from pathlib import Path
import shutil
import uno
from com.sun.star.document import MacroExecMode

from ooodev.calc import CalcDoc, ZoomKind
from ooodev.dialog.msgbox import (
    MessageBoxType,
    MessageBoxButtonsEnum,
    MessageBoxResultsEnum,
)
from ooodev.loader import Lo

# import from this example custom library
# from odev_context_link_lib.dispatch_provider_interceptor import (
#     DispatchProviderInterceptor,
# )
import macro_code


# def register_interceptor(doc: CalcDoc):
#     inst = DispatchProviderInterceptor()  # singleton
#     frame = doc.get_frame()
#     frame.registerDispatchProviderInterceptor(inst)


def main() -> int:
    _ = Lo.load_office(Lo.ConnectSocket())
    try:
        fnm = Path(__file__).parent / "data" / "src_doc" / "links_only.ods"
        tmp_dir = Lo.tmp_dir / "links"
        tmp_dir.mkdir(parents=True, exist_ok=True)
        dest_file = tmp_dir / fnm.name
        shutil.copy(fnm, dest_file)

        doc = CalcDoc.open_doc(
            fnm=dest_file,
            visible=True,
            MacroExecutionMode=MacroExecMode.ALWAYS_EXECUTE_NO_WARN,
        )
        # wait for document to be ready
        Lo.delay(300)
        doc.zoom(ZoomKind.ZOOM_100_PERCENT)
        macro_code.register_url_interceptor()

        # a breakpoint can be set below to inspect the menu.

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
            # un-register or expect a crash
            macro_code.unregister_url_interceptor()
            print("Keeping document open")
    except Exception:
        Lo.close_office()
        raise
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
