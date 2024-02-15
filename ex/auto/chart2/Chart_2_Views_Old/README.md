# Demonstrates creating charts

![charts_happy](https://user-images.githubusercontent.com/4193389/198873533-36de5d26-1071-467b-95f4-2e557b4017cb.png)

Demonstrates loading a spread sheet into Calc and dynamically inserting charts.
There are a total of 17 different charts that can be dynamically created by this demo.

A message box is display once the document has been created asking if you want to close the document.

## UPDATE

Since [OOO Development Tools] version `0.27.0` the there is another more flexible way to create charts.
A new version of the original `Chart_Views` example has been created. This example still works but is now renamed to `Chart_2_Views_Old`.
The new example is located in the [Chart_2_Views](../Chart_2_Views/) folder.

## NOTE

There is currently a [bug](https://bugs.documentfoundation.org/show_bug.cgi?id=151846) in LibreOffice `7.4` that does not allow the `Chart2` class to load.
The `Chart2` has been tested with LibreOffice `7.3` and `7.5+`.

## Options

The type of chart created is determined by the `-k` option.

Possible `-k` options are:

- area
- bar
- bubble_labeled
- col
- col_line
- col_multi
- donut
- happy_stock
- line
- lines
- net
- pie
- pie_3d
- scatter
- scatter_line_error
- scatter_line_log
- stock_prices

## Automate

### Dev Container

From current example folder.

```sh
python -m start -k happy_stock
```

### Cross Platform

From current example folder.

```sh
python -m start -k happy_stock
```

### Linux/Mac

```sh
python ./ex/auto/chart2/Chart_2_Views/start.py -k happy_stock
```

### Windows

```ps
python .\ex\auto\chart2\Chart_2_Views\start.py -k happy_stock
```

## Live LibreOffice Python

Instructions to run this example in [Live-LibreOffice-Python](https://github.com/Amourspirit/live-libreoffice-python).

Start Live-LibreOffice-Python in a Codespace or in a Dev Container.

In the terminal run:

```bash
cd examples
gitget 'https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/chart2/Chart_2_Views'
```

This will copy the `Chart_2_Views` example to the examples folder.

In the terminal run:

```bash
cd Chart_2_Views
python -m start
```

[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/