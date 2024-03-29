# Linear Solve

<p align="center">
<img src="https://user-images.githubusercontent.com/4193389/205754908-519f4c4d-c74c-41c4-82b0-a3211cd82bc4.png" width="960" height="360">
</p>

Examples of Linear Solve in a spreadsheet.

This demo uses This demo uses [OOO Development Tools] (OooDev).

See Also:

- [OOO Development Tools - Chapter 27. Functions and Data Analysis](https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part4/chapter27.html)
- [NonLinear Solve](./odev_non_linear_solve/)

## Automate

### Dev Container

From this folder.

```sh
python -m start
```

### Cross Platform

From this folder.

```sh
python -m start
```

### Linux/Mac

```sh
python ./ex/auto/calc/odev_linear_solve/start.py
```

### Windows

```ps
python .\ex\auto\calc\odev_linear_solve\start.py
```

## Live LibreOffice Python

Instructions to run this example in [Live-LibreOffice-Python](https://github.com/Amourspirit/live-libreoffice-python).

Start Live-LibreOffice-Python in a Codespace or in a Dev Container.

In the terminal run:

```bash
cd examples
gitget 'https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/calc/odev_linear_solve'
```

This will copy the `odev_linear_solve` example to the examples folder.

In the terminal run:

```bash
cd odev_linear_solve
python -m start
```

## Output

```text
Services offered by the solver:
  com.sun.star.comp.Calc.CoinMPSolver
  com.sun.star.comp.Calc.LpsolveSolver
  com.sun.star.comp.Calc.NLPSolver.DEPSSolverImpl
  com.sun.star.comp.Calc.NLPSolver.SCOSolverImpl 
  com.sun.star.comp.Calc.SwarmSolver

Solver result: 
  B3 == 6315.6250 
Solver variables: 
  B1 == 21.8750   
  B2 == 53.1250 
```

[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/
