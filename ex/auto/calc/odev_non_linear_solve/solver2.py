from __future__ import annotations

import uno
from com.sun.star.sheet import XSolver

from ooodev.exceptions import ex
from ooodev.calc import Calc
from ooodev.calc import CalcDoc
from ooodev.loader import Lo
from ooodev.utils.props import Props


class Solver2:
    @staticmethod
    def main(verbose: bool = False) -> None:
        with Lo.Loader(
            connector=Lo.ConnectPipe(), opt=Lo.Options(verbose=verbose)
        ) as loader:
            doc = CalcDoc.create_doc(loader)
            sheet = doc.sheets[0]

            # specify the variable cells
            cell = sheet["B1"]
            x_pos = cell.cell_obj.get_cell_address()
            cell = cell.get_cell_down()  # B2
            y_pos = cell.cell_obj.get_cell_address()
            vars = (x_pos, y_pos)

            # specify profit equation
            cell = cell.get_cell_down()  # B3
            cell.value = "=B1*B2"
            objective = cell.cell_obj.get_cell_address()

            # set up equation formula without inequality (only one needed)
            # x^2 + y^2
            cell = cell.get_cell_down()  # B4
            cell.value = "=B1*B1 + B2*B2"

            # create three constraints (using the 3 variables)

            sc1 = cell.make_constraint(num=1, op=">=")
            #   x^2 + y^2 >= 1
            sc2 = cell.make_constraint(num=2, op="<=")
            #   x^2 + y^2 <= 2

            constraints = (sc1, sc2)

            # initialize the nonlinear solver (SCO)
            try:
                solver = doc.lo_inst.create_instance_mcf(
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

            Props.show_obj_props("Solver", solver)
            # switch off nonlinear dialog about current progress
            # and restrict the search to the top-right quadrant of the graph
            Props.set(solver, EnhancedSolverStatus=False, AssumeNonNegative=True)

            # execute the solver
            solver.solve()
            Calc.solver_report(solver)
            doc.close_doc()
