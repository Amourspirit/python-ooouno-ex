from __future__ import annotations

import uno
from com.sun.star.sheet import XSolver

from ooodev.office.calc import Calc
from ooodev.utils.lo import Lo
from ooodev.utils.props import Props

from ooo.dyn.sheet.solver_constraint_operator import SolverConstraintOperator

class LinearSolve:
    def main(self) -> None:
        with Lo.Loader(connector=Lo.ConnectPipe()) as loader:
            doc = Calc.create_doc(loader)
            sheet = Calc.get_sheet(doc=doc)
            Calc.list_solvers()

            # specify the variable cells
            xpos = Calc.get_cell_address(sheet=sheet, cell_name="B1")  # X
            ypos = Calc.get_cell_address(sheet=sheet, cell_name="B2")  # Y

            vars = (xpos, ypos)

            # specify profit equation
            Calc.set_val(value="=143*B1 + 60*B2", sheet=sheet, cell_name="B3")
            profit_eq = Calc.get_cell_address(sheet, "B3")

            # set up equation formulae without inequalities
            Calc.set_val(value="=120*B1 + 210*B2", sheet=sheet, cell_name="B4")
            Calc.set_val(value="=110*B1 + 30*B2", sheet=sheet, cell_name="B5")
            Calc.set_val(value="=B1 + B2", sheet=sheet, cell_name="B6")

            # create the constraints
            # constraints are equations and their inequalities
            sc1 = Calc.make_constraint(num=15000, op="<=", sheet=sheet, cell_name="B4")
            #   20x + 210y <= 15000
            #   B4 is the address of the cell that is constrained
            sc2 = Calc.make_constraint(num=4000, op=SolverConstraintOperator.LESS_EQUAL, sheet=sheet, cell_name="B5")
            #   110x + 30y <= 4000
            sc3 = Calc.make_constraint(num=75, op="<=", sheet=sheet, cell_name="B6")
            #   x + y <= 75

            # could also include x >= 0 and y >= 0
            constraints = (sc1, sc2, sc3)
            
            # for unknown reason CoinMPSolver stopped working on linux.
            # Ubuntu 22.04 LibreOffice 7.3 no-longer list com.sun.star.comp.Calc.CoinMPSolver
            # as a reported service.
            # strangly Windows 10, LibreOffice 7.3 does still list com.sun.star.comp.Calc.CoinMPSolver
            # as a service.
            # solver = "com.sun.star.comp.Calc.LpsolveSolver"
            solver = "com.sun.star.comp.Calc.CoinMPSolver"

            # initialize the linear solver (CoinMP or basic linear)
            solver = Lo.create_instance_mcf(XSolver, solver, raise_err=True)
            solver.Document = doc
            solver.Objective = profit_eq
            solver.Variables = vars
            solver.Constraints = constraints
            solver.Maximize = True

            # restrict the search to the top-right quadrant of the graph
            Props.set(solver, NonNegative=True)

            # execute the solver
            solver.solve()
            Calc.solver_report(solver)
            Lo.close_doc(doc)