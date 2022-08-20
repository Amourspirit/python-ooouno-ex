#!/usr/bin/env python
# coding: utf-8
import random

from ooodev.office.write import Write
from ooodev.utils.date_time_util import DateUtil
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo


def main() -> int:

    delay = 2_000  # delay so users can see changes.

    with Lo.Loader(Lo.ConnectSocket()) as loader:

        doc = Write.create_doc(loader=loader)

        try:
            GUI.set_visible(is_visible=True, odoc=doc)

            cursor = Write.get_cursor(doc)
            Write.append_para(cursor, "Math Questions")
            Write.style_prev_paragraph(cursor, "Heading 1")

            Write.append_para(cursor, "Solve the following formulae for x:\n")

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

                    # formulas should be wrapped in {} but for fromatting reasons it is easier to work with [] and replace later.
                    if choice == 0:
                        formula = f"[[[sqrt[{iA}x]] over {iB}] + [{iC} over {iD}]=[{iE} over {iF1} ]]"
                    elif choice == 1:
                        formula = f"[[[{iA}x] over {iB}] + [{iC} over {iD}]=[{iE} over {iF1}]]"
                    else:
                        formula = f"[{iA}x + {iB} = {iC}]"

                    # replace [] with {}
                    Write.add_formula(cursor, formula.replace("[", "{").replace("]", "}"))
                    Write.end_paragraph(cursor)

            Write.append_para(cursor, f"Timestamp: {DateUtil.time_stamp()}")

            Lo.delay(delay)
            Lo.save_doc(doc, "mathQuestions.pdf")

        finally:
            Lo.close_doc(doc)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
