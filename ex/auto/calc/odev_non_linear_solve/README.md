# NonLinear Solve

<p align="center">
<img src="https://user-images.githubusercontent.com/4193389/205730756-39fd9a7b-65ef-4a8d-a67a-4b3ab0580235.png" width="414" height="276">
</p>

Example of Non-Linear Solve in a spreadsheet.

Two different examples are included. Include parameter `-s 1` for the first example and
`-s 2` for the second example.

This demo uses This demo uses [OOO Development Tools] (ODEV).

See Also:

- [OOO Development Tools - Chapter 27. Functions and Data Analysis](https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part4/chapter27.html)
- [Linear Solve](./odev_linear_solve/)

## Automate

### Command Line Parameters

- `-s <num>` runs demo for pivot table example one (`-s 1`) or two (`-s 2`).
- `-h` Displays help on options.


### Cross Platform

From this folder.

```sh
python -m start -s 1
```

### Linux/Mac

```sh
python ./ex/auto/calc/odev_non_linear_solve/start.py -s 1
```

### Windows

```ps
python .\ex\auto\calc\odev_non_linear_solve\start.py -s 1
```

## Output

### Solver1

```text
Solver Properties
  AssumeNonNegative: False
  SwarmSize: 70
  LearningCycles: 2000
  GuessVariableRange: True
  VariableRangeThreshold: 3.0
  UseACRComparator: False
  UseRandomStartingPoint: False
  UseStrongerPRNG: False
  StagnationLimit: 70
  Tolerance: 1e-06
  EnhancedSolverStatus: True
  LibrarySize: 210

Solver result: 
  B4 == 10.0000
Solver variables:
  B1 == 6.0000
  B2 == 8.0000
  B3 == 4.0000
```

### Solver2

```text
Solver Properties
  AssumeNonNegative: False
  SwarmSize: 70
  LearningCycles: 2000
  GuessVariableRange: True
  VariableRangeThreshold: 3.0
  UseACRComparator: False
  UseRandomStartingPoint: False
  UseStrongerPRNG: False
  StagnationLimit: 70
  Tolerance: 1e-06
  EnhancedSolverStatus: True
  LibrarySize: 210

Solver result: 
  B3 == 2.0000
Solver variables:
  B1 == 1.0001
  B2 == 0.9999
```

[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/
