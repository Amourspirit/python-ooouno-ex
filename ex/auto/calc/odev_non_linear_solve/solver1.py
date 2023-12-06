from __future__ import annotations

import uno
from com.sun.star.sheet import XSolver

from ooodev.exceptions import ex
from ooodev.calc import Calc
from ooodev.calc import CalcDoc
from ooodev.utils.lo import Lo
from ooodev.utils.props import Props


class Solver1:
    @staticmethod
    def main(verbose: bool = False) -> None:
        with Lo.Loader(
            connector=Lo.ConnectPipe(), opt=Lo.Options(verbose=verbose)
        ) as loader:
            doc = CalcDoc(Calc.create_doc(loader))
            sheet = doc.get_sheet(0)

            # specify the variable cells
            x_pos = sheet.get_cell_address(cell_name="B1")  # X
            y_pos = sheet.get_cell_address(cell_name="B2")  # Y
            z_pos = sheet.get_cell_address(cell_name="B3")  # z
            vars = (x_pos, y_pos, z_pos)

            # set up equation formula without inequality
            sheet.set_val(value="=B1+B2-B3", cell_name="B4")
            objective = sheet.get_cell_address(cell_name="B4")

            # create three constraints (using the 3 variables)

            sc1 = sheet.make_constraint(num=6, op="<=", cell_name="B1")
            #   x <= 6
            sc2 = sheet.make_constraint(num=8, op="<=", cell_name="B2")
            #   y <= 8
            sc3 = sheet.make_constraint(num=4, op=">=", cell_name="B3")
            #   z >= 4

            constraints = (sc1, sc2, sc3)

            # initialize the nonlinear solver (SCO)
            try:
                solver = Lo.create_instance_mcf(
                    XSolver,
                    "com.sun.star.comp.Calc.NLPSolver.SCOSolverImpl",
                    raise_err=True,
                )
            except ex.CreateInstanceMcfError:
                print(
                    'Solver not available on this system: "com.sun.star.comp.Calc.NLPSolver.SCOSolverImpl"'
                )
                doc.close_doc()
                return
            solver.Document = doc.component
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
            doc.close_doc()
