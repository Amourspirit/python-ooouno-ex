# Pivot Tables

<p align="center">
<img src="https://user-images.githubusercontent.com/4193389/205518229-e59beb75-21c0-44f4-bde0-b8665a43afb0.png" width="567" height="346">
</p>

Example create a new sheet containing pivot table in the document.

Two different examples are included. Include parameter `-p 1` for the first example and
`-p 2` for the second example.

This demo uses This demo uses [OOO Development Tools] (ODEV).

See Also:

- [OOO Development Tools - Chapter 27. Functions and Data Analysis](https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part4/chapter27.html)

## Automate

A message box is display once the document has been processed asking if you want to close the document.

### Command Line Parameters

- `-p <num>` runs demo for pivot table example one (`-p 1`) or two (`-p 2`).
- `-o` Optional output path such as `-o "tmp/piviot.ods"`
- `-h` Displays help on options.

### Dev Container

Run from this folder.

```sh
python -m start -p 1
```

### Cross Platform

From this folder.

```sh
python -m start -p 1
```

### Linux/Mac

```sh
python ./ex/auto/calc/odev_pivot_table/start.py -p 1
```

### Windows

```ps
python .\ex\auto\calc\odev_pivot_table\start.py -p 1
```

## Live LibreOffice Python

Instructions to run this example in [Live-LibreOffice-Python](https://github.com/Amourspirit/live-libreoffice-python).

Start Live-LibreOffice-Python in a Codespace or in a Dev Container.

In the terminal run:

```bash
cd examples
gitget 'https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/calc/odev_pivot_table'
```

This will copy the `odev_pivot_table` example to the examples folder.

In the terminal run:

```bash
cd odev_pivot_table
python -m start -h
```

[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/
