from __future__ import annotations

import uno
from com.sun.star.sheet import XGoalSeek

from ooodev.exceptions.ex import GoalDivergenceError
from ooodev.office.calc import Calc
from ooodev.utils.lo import Lo


class GoalSeek:
    def main(self) -> None:
        with Lo.Loader(connector=Lo.ConnectPipe()) as loader:
            doc = Calc.create_doc(loader)
            sheet = Calc.get_sheet(doc=doc)
            gs = Lo.qi(XGoalSeek, doc)

            # -------------------------------------------------
            # x-variable and starting value
            Calc.set_val(value=9, sheet=sheet, cell_name="C1")
            # formula
            Calc.set_val(value="=SQRT(C1)", sheet=sheet, cell_name="C2")
            x = Calc.goal_seek(gs=gs, sheet=sheet, cell_name="C1", formula_cell_name="C2", result=4.0)
            print(f"x == {x}\n")  # 16.0

            # -------------------------------------------------
            try:
                x = Calc.goal_seek(gs=gs, sheet=sheet, cell_name="C1", formula_cell_name="C2", result=-4.0)
                # The formula is still y = sqrt(x)
                # Find x when sqrt(x) == -4, which is impossible
                print(f"x == {x} when sqrt(x) == -4\n")
                
            except GoalDivergenceError as e:
                print(e)
                print()
            # -------------------------------------------------
            # x-variable and starting value
            Calc.set_val(sheet=sheet, cell_name="D1", value=0.8)
            # formula
            Calc.set_val(sheet=sheet, cell_name="D2", value="=(D1^2 - 1)/(D1 - 1)")
            # The formula is y = (x^2 -1)/(x-1)
            # After factoring, this is just y = x+1
            x = Calc.goal_seek(gs=gs, sheet=sheet, cell_name="D1", formula_cell_name="D2", result=2)
            print(f"x == {x} when x+1 == 2\n")

            # -------------------------------------------------
            #  x-variable; starting capital
            Calc.set_val(value=100000, sheet=sheet, cell_name="B1")
            # n, no. of years
            Calc.set_val(value=1, sheet=sheet, cell_name="B2")
            # i, interest rate (7.5%)
            Calc.set_val(value=0.075, sheet=sheet, cell_name="B3")
            # formula
            Calc.set_val("=B1*B2*B3", sheet, "B4")
            # The formula is Annual interest = x*n*r
            # where capital (x), number of years (n), and interest rate (r).
            # Find the capital, if the other values are given.
            x = Calc.goal_seek(gs=gs, sheet=sheet, cell_name="B1", formula_cell_name="B4", result=15000)
            # x is 200,000
            print(
                (
                    f"x == {x} when x*"
                    f'{Calc.get_val(sheet=sheet, cell_name="B2")}*'
                    f'{Calc.get_val(sheet=sheet, cell_name="B3")}'
                    " == 15000\n"
                )
            )

            # -------------------------------------------------
            # x-variable and starting value
            Calc.set_val(value=0, sheet=sheet, cell_name="E1")
            # formula
            Calc.set_val(value="=(E1^3 - 2*E1 + 2", sheet=sheet, cell_name="E2")
            x = Calc.goal_seek(gs=gs, sheet=sheet, cell_name="E1", formula_cell_name="E2", result=0)
            # x is -1.7692923428381226 so not using Newton's method which oscillates between 0 and 1
            print(f"x == {x} when formula == 0\n")
