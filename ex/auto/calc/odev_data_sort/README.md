# Data Sort

<p align="center">
<img src="https://user-images.githubusercontent.com/4193389/204033934-8585c854-7203-41fa-a0cc-7b9232cc700a.png" width="319" height="240">
</p>

Example of adding data and sorting data in a Spreadsheet

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
python ./ex/auto/calc/odev_data_sort/start.py
```

### Windows

```ps
python .\ex\auto\calc\odev_data_sort\start.py
```

## Example Usage

Starts LibreOffice Calc, Builds spreadsheet tables, adds an image, inserts chart and saves output to `tmp/Build.ods`

```ps
python .\ex\auto\calc\odev_data_sort\start.py -o "tmp/dataSort.ods"
```

[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/
