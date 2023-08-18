# Show Sheet

<p align="center">
<img src="https://user-images.githubusercontent.com/4193389/194169727-f5a61ab2-e336-42c3-8ef1-31299b81100d.jpg" width="348" height="232">
</p>

Example of opening a spreadsheet and inputting a password (`foobar`) to unlock sheet.

Also demonstrates how to create input password dialog and message dialog.

Optionally saves the input file as a new file.

This demo uses This demo uses [OOO Development Tools] (OooDev).

See Also:

- [OOO Development Tools - Chapter 20. Spreadsheet Displaying and Creation](https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part4/chapter20.html)


## Automate

A message box is display once the document has been processed asking if you want to close the document.

The following command will run automation that opens Calc document and ask for password.

### Dev Container

Run from this folder.

```sh
python -m start --show --file "./data/totals.ods"
```

### Cross Platform

From this folder.

```sh
python -m start --show --file "./data/totals.ods"
```

### Linux/Mac

```sh
python ./ex/auto/calc/odev_show_sheet/start.py --show --file "ex/auto/calc/odev_show_sheet/data/totals.ods" --out "tmp/totals.pdf"
```

Alternatively

```sh
python ./ex/auto/calc/odev_show_sheet/start.py --show --file "ex/auto/calc/odev_show_sheet/data/sorted.csv" --out "tmp/totals.html"
```

### Windows

```ps
python .\ex\auto\calc\odev_show_sheet\start.py --show --file "ex/auto/calc/odev_show_sheet/data/totals.ods" --out "tmp/totals.pdf"
```

Alternatively

```ps
python .\ex\auto\calc\odev_show_sheet\start.py --show --file "ex/auto/calc/odev_show_sheet/data/sorted.csv" --out "tmp/totals.html"
```

## Live LibreOffice Python

Instructions to run this example in [Live-LibreOffice-Python](https://github.com/Amourspirit/live-libreoffice-python).

Start Live-LibreOffice-Python in a Codespace or in a Dev Container.

In the terminal run:

```bash
cd examples
gitget 'https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/calc/odev_show_sheet'
```

This will copy the `odev_show_sheet` example to the examples folder.

In the terminal run:

```bash
cd odev_show_sheet
python -m start -h
```

[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/
