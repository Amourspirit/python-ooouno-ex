from __future__ import annotations
import random
from pathlib import Path

from ooodev.dialog.msgbox import (
    MsgBox,
    MessageBoxType,
    MessageBoxButtonsEnum,
    MessageBoxResultsEnum,
)
from ooodev.write import Write, WriteDoc
from ooodev.utils.date_time_util import DateUtil
from ooodev.utils.lo import Lo


def main() -> int:
    delay = 2_000  # delay so users can see changes.

    loader = Lo.load_office(Lo.ConnectPipe())

    doc = WriteDoc(Write.create_doc(loader=loader))

    try:
        doc.set_visible()

        cursor = doc.get_cursor()
        cursor.append_para("Math Questions")
        cursor.style_prev_paragraph("Heading 1")

        cursor.append_para("Solve the following formulae for x:\n")

        # lock screen updating and add formulas
        # locking screen is not strictly necessary but is faster when add lost of input.
        with Lo.ControllerLock():
            for _ in range(10):  # generate 10 random formulae
                iA = random.randint(0, 7) + 2
                iB = random.randint(0, 7) + 2
                iC = random.randint(0, 8) + 1
                iD = random.randint(0, 7) + 2
                iE = random.randint(0, 8) + 1
                iF1 = random.randint(0, 7) + 2

                choice = random.randint(0, 2)

                # formulas should be wrapped in {} but for formatting reasons it is easier to work with [] and replace later.
                if choice == 0:
                    formula = f"[[[sqrt[{iA}x]] over {iB}] + [{iC} over {iD}]=[{iE} over {iF1} ]]"
                elif choice == 1:
                    formula = (
                        f"[[[{iA}x] over {iB}] + [{iC} over {iD}]=[{iE} over {iF1}]]"
                    )
                else:
                    formula = f"[{iA}x + {iB} = {iC}]"

                # replace [] with {}
                cursor.add_formula(formula.replace("[", "{").replace("]", "}"))
                cursor.end_paragraph()

        cursor.append_para(f"Timestamp: {DateUtil.time_stamp()}")

        Lo.delay(delay)
        msg_result = MsgBox.msgbox(
            "Do you wish to save document?",
            "Save",
            boxtype=MessageBoxType.QUERYBOX,
            buttons=MessageBoxButtonsEnum.BUTTONS_YES_NO,
        )
        if msg_result == MessageBoxResultsEnum.YES:
            pth = Path.cwd() / "tmp"
            pth.mkdir(exist_ok=True)
            doc.save_doc(pth / "mathQuestions.pdf")

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
