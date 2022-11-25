# Filler

<p align="center">
<img src="https://user-images.githubusercontent.com/4193389/204039084-db5f7e1b-9aab-4525-875b-dbdab22b7b13.png" width="247" height="247">
</p>

Example of using a fill series in a spreadsheet.

This demo uses This demo uses [OOO Development Tools] (ODEV).

See Also:

- [OOO Development Tools - Chapter 24. Complex Data Manipulation](https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part4/chapter24.html)

## Automate

A message box is display once the document has been processed asking if you want to close the document.

### Cross Platform

From this folder.

```sh
python -m start
```

### Linux/Mac

```sh
python ./ex/auto/calc/odev_filler/start.py
```

### Windows

```ps
python .\ex\auto\calc\odev_filler\start.py
```

## Example Usage

Starts LibreOffice Calc, runs automation and saves output to `tmp/Filler.ods`

```ps
python .\ex\auto\calc\odev_filler\start.py -o "tmp/Filler.ods"
```

[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/
