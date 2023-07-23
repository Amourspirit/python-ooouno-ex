# Calc Functions

<p align="center">
<img src="https://user-images.githubusercontent.com/4193389/205467747-3de8424a-ba2c-4613-abff-d901cc2074cf.png">
</p>


Demostrates:

Use/list/find Calc functions.

1. Database: for extracting information from Calc tables, where the data is organized into rows.
The “Database” name is a little misleading, but the documentation makes the point that
Calc database functions have nothing to do with Base databases. Chapter 13 of the
Calc User Guide (“Calc as a Simple Database”) explains the distinction in detail.
Date and Time.
2. Financial: for business calculations;
3. Information: many of these return boolean information about cells, such as whether a cell contains text or a formula;
4. Logical: functions for boolean logic;
5. Mathematical: trigonometric, hyperbolic, logarithmic, and summation functions; e.g. see ROUND, SIN, and RADIANS below;
6. Array: many of these operations treat cell ranges like 2D arrays; i.e. see TRANSPOSE below;
7. Statistical: for statistical and probability calculations; i.e., see AVERAGE and SLOPE below;
8. Spreadsheet: for finding values in tables, cell ranges, and cells;
9. Text: string manipulation functions;
10.  Add-ins: a catch-all category that includes a lot of functions – extra data and time operations, conversion functions between number bases, more statistics, and complex numbers.
11. See IMSUM and ROMAN below for examples.
The “Add-ins” documentation starts at [Calc Add-in Functions](https://help.libreoffice.org/latest/en-US/text/scalc/01/04060111.html), and continues in
[Add-in Functions, List of Analysis Functions Part One](https://help.libreoffice.org/latest/en-US/text/scalc/01/04060115.html) and [Add-in Functions, List of Analysis Functions Part Two](https://help.libreoffice.org/latest/en-US/text/scalc/01/04060116.html).

This demo uses This demo uses [OOO Development Tools] (ODEV).

See Also:

- [OOO Development Tools - Chapter 27. Functions and Data Analysis](https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part4/chapter27.html)

## Automate

### Cross Platform

From this folder.

```sh
python -m start
```

### Linux/Mac

```sh
python ./ex/auto/calc/odev_functions/start.py
```


### Windows

```ps
python .\ex\auto\calc\odev_functions\start.py
```

## Live LibreOffice Python

Instructions to run this example in [Live-LibreOffice-Python](https://github.com/Amourspirit/live-libreoffice-python).

Start Live-LibreOffice-Python in a Codespace or in a Dev Container.

In the terminal run:

```bash
cd examples
gitget 'https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/calc/odev_functions'
```

This will copy the `odev_functions` example to the examples folder.

In the terminal run:

```bash
cd odev_functions
python -m start
```

## Output

```text
ROUND result for 1.999 is: 2.0

SIN result for 30 degrees is:0.500
Average of the numbers is: 3.0    

SLOPE of the line: 3.0

ZTEST result for data ((1,2,3),) and 2.0 is: 0.5
Easter Sunday (d/m/y): 5/4/2015

Row x Column size: 3 x 3
  1.0   4.0
  2.0   5.0
  3.0   6.0

13+4j + 5+3j: 18+7j

100 to hex: 0064

ROT13 of "hello": uryyb

999 in Roman numerals: CMXCIX or IM

Relative address for (5,2): E2

Properties for "EASTERSUNDAY"":
  Id: 380
  Category: 2
  Name: EASTERSUNDAY
  Description: Calculates the date of Easter Sunday in a given year.
  Arguments: [Year]

No. of arguments: 1
1. Argument name: Year
  Description: 'An integer between 1583 and 9956, or 0 and 99 (19xx or 20xx depending on the option set).'
  Is optional?: False


Properties for "ROMAN"":
  Id: 383
  Category: 10
  Name: ROMAN
  Description: Converts a number to a Roman numeral.
  Arguments: [Number, Mode (optional)]

No. of arguments: 2
1. Argument name: Number
  Description: 'The number to be converted to a Roman numeral must be in the 0 - 3999 range.'
  Is optional?: False

2. Argument name: Mode
  Description: 'The more this value increases, the more the Roman numeral is simplified. The value must be in the 0 - 4 range.'
  Is optional?: True

Recently used functions 5
  SKEW
  MEDIAN
  STANDARDIZE
  AVEDEV
  PI
```

[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/
