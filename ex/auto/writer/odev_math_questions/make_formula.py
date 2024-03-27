from __future__ import annotations, unicode_literals
import random
from ooodev.write import WriteDoc
from ooodev.utils.date_time_util import DateUtil


def generate_formula(*args) -> None:
    doc = WriteDoc.from_current_doc()
    cursor = doc.get_cursor()
    cursor.append_para("Math Questions")
    cursor.style_prev_paragraph("Heading 1")
    cursor.append_para("Solve the following formulae for x:\n")

    # lock screen updating and add formulas
    # locking screen is not strictly necessary but is faster when add lost of input.
    with doc:
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
                formula = (
                    f"[[[sqrt[{iA}x]] over {iB}] + [{iC} over {iD}]=[{iE} over {iF1} ]]"
                )
            elif choice == 1:
                formula = f"[[[{iA}x] over {iB}] + [{iC} over {iD}]=[{iE} over {iF1}]]"
            else:
                formula = f"[{iA}x + {iB} = {iC}]"

            # replace [] with {}
            cursor.add_formula(formula.replace("[", "{").replace("]", "}"))
            cursor.end_paragraph()

        cursor.append_para(f"Timestamp: {DateUtil.time_stamp()}")


g_exportedScripts = (generate_formula,)
