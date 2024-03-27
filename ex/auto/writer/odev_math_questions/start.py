from __future__ import annotations
from pathlib import Path

from ooodev.dialog.msgbox import (
    MessageBoxType,
    MessageBoxButtonsEnum,
    MessageBoxResultsEnum,
)
from ooodev.loader import Lo
from ooodev.write import WriteDoc
from make_formula import generate_formula


def main() -> int:
    delay = 2_000  # delay so users can see changes.

    loader = Lo.load_office(Lo.ConnectPipe())

    doc = WriteDoc.create_doc(loader=loader, visible=True)

    try:
        generate_formula()

        Lo.delay(delay)
        msg_result = doc.msgbox(
            "Do you wish to save document?",
            "Save",
            boxtype=MessageBoxType.QUERYBOX,
            buttons=MessageBoxButtonsEnum.BUTTONS_YES_NO,
        )
        if msg_result == MessageBoxResultsEnum.YES:
            pth = Path.cwd() / "tmp"
            pth.mkdir(exist_ok=True)
            doc.save_doc(pth / "mathQuestions.pdf")

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

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
