from __future__ import annotations

import uno
from com.sun.star.sheet import XSolver

from ooodev.exceptions import ex
from ooodev.calc import Calc
from ooodev.calc import CalcDoc
from ooodev.loader import Lo
from ooodev.utils.props import Props


class Solver1:
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
            cell = cell.get_cell_down()  # B3
            z_pos = cell.cell_obj.get_cell_address()
            vars = (x_pos, y_pos, z_pos)

            # set up equation formula without inequality
            cell = cell.get_cell_down()  # B4
            cell.value = "=B1+B2-B3"
            objective = cell.cell_obj.get_cell_address()

            # create three constraints (using the 3 variables)

            cell = sheet["B1"]
            sc1 = cell.make_constraint(num=6, op="<=")
            #   x <= 6
            cell = cell.get_cell_down()  # B2
            sc2 = cell.make_constraint(num=8, op="<=")
            #   y <= 8
            cell = cell.get_cell_down()  # B3
            sc3 = cell.make_constraint(num=4, op=">=")
            #   z >= 4

            constraints = (sc1, sc2, sc3)

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

            # restrict the search to the top-right quadrant of the graph
            Props.show_obj_props("Solver", solver)
            # switch off nonlinear dialog about current progress
            Props.set(solver, EnhancedSolverStatus=False)

            # execute the solver
            solver.solve()
            # Profit max == 10; vars are very close to 6, 8, and 4, but off by 6-7 dps
            Calc.solver_report(solver)
            doc.close_doc()
