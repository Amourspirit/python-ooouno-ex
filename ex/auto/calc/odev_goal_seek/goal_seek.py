from __future__ import annotations

import uno
from com.sun.star.sheet import XGoalSeek

from ooodev.exceptions.ex import GoalDivergenceError
from ooodev.calc import CalcDoc
from ooodev.loader import Lo


class GoalSeek:
    def main(self) -> None:
        with Lo.Loader(connector=Lo.ConnectPipe()) as loader:
            doc = CalcDoc.create_doc(loader)
            sheet = doc.sheets[0]
            gs = doc.qi(XGoalSeek, True)

            # -------------------------------------------------
            # x-variable and starting value
            cell1 = sheet["C1"]
            cell1.value = 9
            # formula
            cell2 = cell1.get_cell_down()
            cell2.value = "=SQRT(C1)"

            x = cell1.goal_seek(gs=gs, formula_cell_name=cell2.cell_obj, result=4.0)
            print(f"x == {x}\n")  # 16.0

            # -------------------------------------------------
            try:
                x = cell1.goal_seek(
                    gs=gs,
                    formula_cell_name=cell2.cell_obj,
                    result=-4.0,
                )
                # The formula is still y = sqrt(x)
                # Find x when sqrt(x) == -4, which is impossible
                print(f"x == {x} when sqrt(x) == -4\n")

            except GoalDivergenceError as e:
                print(e)
                print()
            # -------------------------------------------------
            # x-variable and starting value
            cell1 = cell1.get_cell_right()  # D1
            cell2 = cell2.get_cell_down()  # D2

            cell1.value = 0.8
            # formula
            cell2.value = "=(D1^2 - 1)/(D1 - 1)"
            # The formula is y = (x^2 -1)/(x-1)
            # After factoring, this is just y = x+1
            x = cell1.goal_seek(gs=gs, formula_cell_name=cell2.cell_obj, result=2)
            print(f"x == {x} when x+1 == 2\n")

            # -------------------------------------------------
            #  x-variable; starting capital
            cell1 = sheet["B1"]
            cell2 = cell1.get_cell_down()  # B2
            cell3 = cell2.get_cell_down()  # B3
            cell4 = cell3.get_cell_down()  # B4

            cell1.value = 100_000
            # n, no. of years
            cell2.value = 1
            # i, interest rate (7.5%)
            cell3.value = 0.075
            # formula
            cell4.value = "=B1*B2*B3"
            # The formula is Annual interest = x*n*r
            # where capital (x), number of years (n), and interest rate (r).
            # Find the capital, if the other values are given.
            x = cell1.goal_seek(gs=gs, formula_cell_name=cell4.cell_obj, result=15000)
            # x is 200,000
            print(
                (
                    f"x == {x} when x*"
                    f"{cell2.get_val()}*"
                    f"{cell3.get_val()}"
                    " == 15000\n"
                )
            )

            # -------------------------------------------------
            # x-variable and starting value
            cell1 = sheet["E1"]
            cell2 = cell1.get_cell_down()
            cell1.value = 0

            # formula
            cell2.value = "=(E1^3 - 2*E1 + 2"
            x = cell1.goal_seek(gs=gs, formula_cell_name=cell2.cell_obj, result=0)
            # x is -1.7692923428381226 so not using Newton's method which oscillates between 0 and 1
            print(f"x == {x} when formula == 0\n")
            doc.close()
