# Garlic Secrets

<p align="center">
<img src="https://user-images.githubusercontent.com/4193389/203686861-4da84607-f65b-479b-9c9d-0321a3733315.png" width="317" height="198">
</p>

This example illustrates how data can be extracted from an existing spreadsheet (`produceSales.xlsx`) using 'general' functions, sheet searching, and sheet range queries. It also has more examples of cell styling, and demonstrates sheet freezing, view pane splitting, pane activation, and the insertion of new rows into a sheet.

This demo uses This demo uses [OOO Development Tools] (OooDev).

See Also:

- [OOO Development Tools - Chapter 23. Garlic Secrets](https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part4/chapter23.html)

## Automate

A message box is display once the document has been processed asking if you want to close the document.

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
python ./ex/auto/calc/odev_garlic_secrets/start.py
```

### Windows

```ps
python .\ex\auto\calc\odev_garlic_secrets\start.py
```

## Live LibreOffice Python

Instructions to run this example in [Live-LibreOffice-Python](https://github.com/Amourspirit/live-libreoffice-python).

Start Live-LibreOffice-Python in a Codespace or in a Dev Container.

In the terminal run:

```bash
cd examples
gitget 'https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/calc/odev_garlic_secrets'
```

This will copy the `odev_garlic_secrets` example to the examples folder.

In the terminal run:

```bash
cd odev_garlic_secrets
python -m start
```

## Example Usage

Starts LibreOffice Calc, runs example and saves output to `tmp/garlic_secrets.ods`

```ps
python .\ex\auto\calc\odev_garlic_secrets\start.py -o "tmp/garlic_secrets.ods"
```

[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/
