# Build Table

<p align="center">
<img src="https://user-images.githubusercontent.com/4193389/202744462-382749d4-1dec-467d-b8c3-88b4dbc3e85f.png" width="486" height="304">
</p>

Example of building different kinds of Spreadsheet Tables.

Also demonstrates create a chart and inserting an image.

This demo uses This demo uses [OOO Development Tools] (OooDev).

See Also:

- [OOO Development Tools - Chapter 20. Spreadsheet Displaying and Creation](https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part4/chapter20.html)
- [OOO Development Tools - Chapter 22. Styles](https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part4/chapter22.html)

## Automate

A message box is display once the document has been processed asking if you want to close the document.

### Dev Container

From this folder.

```sh
python -m start -h
```

### Cross Platform

From this folder.

```sh
python -m start -h
```

### Linux/Mac

```sh
python ./ex/auto/calc/odev_build_table/start.py -h
```

### Windows

```ps
python .\ex\auto\calc\odev_build_table\start.py -h
```

## Live LibreOffice Python

Instructions to run this example in [Live-LibreOffice-Python](https://github.com/Amourspirit/live-libreoffice-python).

Start Live-LibreOffice-Python in a Codespace or in a Dev Container.

In the terminal run:

```bash
cd examples
gitget 'https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/calc/odev_build_table'
```

This will copy the `odev_build_table` example to the examples folder.

In the terminal run:

```bash
cd odev_build_table
python -m start
```

## Example Usage

Starts LibreOffice Calc, Builds spreadsheet tables, adds an image, inserts chart and saves output to `tmp/Build.ods`

```ps
python .\ex\auto\calc\odev_build_table\start.py -c -p -o "tmp/Build.ods"
```

[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/
