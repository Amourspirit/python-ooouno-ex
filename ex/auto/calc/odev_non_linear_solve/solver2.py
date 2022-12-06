from __future__ import annotations

import uno
from com.sun.star.sheet import XSolver

from ooodev.office.calc import Calc
from ooodev.utils.lo import Lo
from ooodev.utils.props import Props


class Solver2:
    @staticmethod
    def main(verose: bool = False) -> None:
        with Lo.Loader(connector=Lo.ConnectPipe(), opt=Lo.Options(verbose=verose)) as loader:
            doc = Calc.create_doc(loader)
            sheet = Calc.get_sheet(doc=doc)

            # specify the variable cells
            xpos = Calc.get_cell_address(sheet=sheet, cell_name="B1")  # X
            ypos = Calc.get_cell_address(sheet=sheet, cell_name="B2")  # Y
            vars = (xpos, ypos)

            # specify profit equation
            Calc.set_val(value="=B1+B2", sheet=sheet, cell_name="B3")
            objective = Calc.get_cell_address(sheet, "B3")

            # set up equation formula without inequality (only one needed)
            # x^2 + y^2
            Calc.set_val(value="=B1*B1 + B2*B2", sheet=sheet, cell_name="B4")

            # create three constraints (using the 3 variables)

            sc1 = Calc.make_constraint(num=1, op=">=", sheet=sheet, cell_name="B4")
            #   x^2 + y^2 >= 1
            sc2 = Calc.make_constraint(num=2, op="<=", sheet=sheet, cell_name="B4")
            #   x^2 + y^2 <= 2

            constraints = (sc1, sc2)

            # initialize the nonlinear solver (SCO)
            solver = Lo.create_instance_mcf(XSolver, "com.sun.star.comp.Calc.NLPSolver.SCOSolverImpl", raise_err=True)
            solver.Document = doc
            solver.Objective = objective
            solver.Variables = vars
            solver.Constraints = constraints
            solver.Maximize = True

            Props.show_obj_props("Solver", solver)
            # switch off nonlinear dialog about current progress
            # and restrict the search to the top-right quadrant of the graph
            Props.set(solver, EnhancedSolverStatus=False, AssumeNonNegative=True)

            # execute the solver
            solver.solve()
            Calc.solver_report(solver)
            Lo.close_doc(doc)
