from __future__ import annotations

import uno
from com.sun.star.sheet import XSolver

from ooo.dyn.sheet.solver_constraint_operator import SolverConstraintOperator

from ooodev.calc import Calc
from ooodev.calc import CalcDoc
from ooodev.loader import Lo
from ooodev.utils.info import Info
from ooodev.utils.props import Props


class LinearSolve:
    @staticmethod
    def main(verbose: bool = False) -> None:
        with Lo.Loader(
            connector=Lo.ConnectPipe(), opt=Lo.Options(verbose=verbose)
        ) as loader:
            doc = CalcDoc.create_doc(loader)
            sheet = doc.get_active_sheet()
            Calc.list_solvers()

            # specify the variable cells
            cell = sheet["B1"]
            x_pos = cell.cell_obj.get_cell_address()  # X
            cell = cell.get_cell_down()  # B2
            y_pos = cell.cell_obj.get_cell_address()  # Y
            vars = (x_pos, y_pos)

            # specify profit equation
            cell = cell.get_cell_down()  # B3
            cell.value = "=143*B1 + 60*B2"
            profit_eq = cell.cell_obj.get_cell_address()

            # set up equation formulae without inequalities
            cell = cell.get_cell_down()  # B4
            cell.value = "=120*B1 + 210*B2"
            cell = cell.get_cell_down()  # B5
            cell.value = "=110*B1 + 30*B2"
            cell = cell.get_cell_down()  # B6
            cell.value = "=B1 + B2"

            # create the constraints
            # constraints are equations and their inequalities
            cell = sheet["B4"]
            sc1 = cell.make_constraint(num=15_000, op="<=")

            #   20x + 210y <= 15000
            #   B4 is the address of the cell that is constrained
            cell = cell.get_cell_down()  # B5
            sc2 = cell.make_constraint(
                num=4_000, op=SolverConstraintOperator.LESS_EQUAL
            )

            #   110x + 30y <= 4000
            cell = cell.get_cell_down()  # B6
            sc3 = cell.make_constraint(num=75, op="<=")
            #   x + y <= 75

            # could also include x >= 0 and y >= 0
            constraints = (sc1, sc2, sc3)

            # for unknown reason CoinMPSolver stopped working on linux.
            # Ubuntu 22.04 LibreOffice 7.3 no-longer list com.sun.star.comp.Calc.CoinMPSolver
            # as a reported service.
            # strangely Windows 10, LibreOffice 7.3 does still list com.sun.star.comp.Calc.CoinMPSolver
            # as a service.
            # srv_solver = "com.sun.star.comp.Calc.LpsolveSolver"
            solvers = Info.get_service_names(service_name="com.sun.star.sheet.Solver")
            potential_solvers = (
                "com.sun.star.comp.Calc.CoinMPSolver",
                "com.sun.star.comp.Calc.LpsolveSolver",
            )

            srv_solver = ""
            for val in potential_solvers:
                if val in solvers:
                    srv_solver = val
                    break

            if not srv_solver:
                raise ValueError("No valid solver was found")
            # initialize the linear solver (CoinMP or basic linear)
            solver = doc.lo_inst.create_instance_mcf(
                XSolver, srv_solver, raise_err=True
            )

            solver.Document = doc.component
            solver.Objective = profit_eq
            solver.Variables = vars
            solver.Constraints = constraints
            solver.Maximize = True

            # restrict the search to the top-right quadrant of the graph
            Props.set(solver, NonNegative=True)

            # execute the solver
            solver.solve()
            Calc.solver_report(solver)
            doc.close_doc()
