from __future__ import annotations

import uno
from com.sun.star.sheet import XSolver

from ooodev.exceptions import ex
from ooodev.office.calc import Calc
from ooodev.utils.lo import Lo
from ooodev.utils.props import Props


class Solver1:
    @staticmethod
    def main(verose: bool = False) -> None:
        with Lo.Loader(connector=Lo.ConnectPipe(), opt=Lo.Options(verbose=verose)) as loader:
            doc = Calc.create_doc(loader)
            sheet = Calc.get_sheet(doc=doc)

            # specify the variable cells
            xpos = Calc.get_cell_address(sheet=sheet, cell_name="B1")  # X
            ypos = Calc.get_cell_address(sheet=sheet, cell_name="B2")  # Y
            zpos = Calc.get_cell_address(sheet=sheet, cell_name="B3")  # z
            vars = (xpos, ypos, zpos)

            # set up equation formula without inequality
            Calc.set_val(value="=B1+B2-B3", sheet=sheet, cell_name="B4")
            objective = Calc.get_cell_address(sheet, "B4")

            # create three constraints (using the 3 variables)

            sc1 = Calc.make_constraint(num=6, op="<=", sheet=sheet, cell_name="B1")
            #   x <= 6
            sc2 = Calc.make_constraint(num=8, op="<=", sheet=sheet, cell_name="B2")
            #   y <= 8
            sc3 = Calc.make_constraint(num=4, op=">=", sheet=sheet, cell_name="B3")
            #   z >= 4

            constraints = (sc1, sc2, sc3)

            # initialize the nonlinear solver (SCO)
            try:
                solver = Lo.create_instance_mcf(XSolver, "com.sun.star.comp.Calc.NLPSolver.SCOSolverImpl", raise_err=True)
            except ex.MissingInterfaceError:
                print('Solver not available on this system: "com.sun.star.comp.Calc.NLPSolver.SCOSolverImpl"')
                Lo.close_doc(doc)
                return
            solver.Document = doc
            solver.Objective = objective
            solver.Variables = vars
            solver.Constraints = constraints
            solver.Maximize = True

            # restrict the search to the top-right quadrant of the graph
            Props.show_obj_props("Solver", solver)
            # switch off nonlinear dialog about current progress
            Props.set(solver, EnhancedSolverStatus=False)

            # execute the solver
            solver.solve()
            # Profit max == 10; vars are very close to 6, 8, and 4, but off by 6-7 dps
            Calc.solver_report(solver)
            Lo.close_doc(doc)
