# Named Ranges

<p align="center">
<img src="https://github.com/Amourspirit/python-ooouno-ex/assets/4193389/3a4773b2-830f-441e-ba7e-5c42402fc13d" width="500" height="498">
</p>

This example illustrates how Named Ranges can be added or removed from a spreadsheet or document. It also demonstrates how to query Named Ranges and how to use them to access data in a spreadsheet.

This demo uses This demo uses [OOO Development Tools] (OooDev).

```python
sheet = doc.sheets[0]
ca = sheet.get_cell_address(cell_name="A2")
sheet_named_rngs = sheet.named_ranges
sheet_named_rngs.add_new_by_name(
    name="my_sheet_range", content="$Sheet.$A$2:$D$5", position=ca
)
```

## Automate

A few message boxes are  display once the document has been processed asking if you want to remove the named ranges and close the document.

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
python ./ex/auto/calc/odev_named_range/start.py
```

### Windows

```ps
python .\ex\auto\calc\odev_named_range\start.py
```

## Live LibreOffice Python

Instructions to run this example in [Live-LibreOffice-Python](https://github.com/Amourspirit/live-libreoffice-python).

Start Live-LibreOffice-Python in a Codespace or in a Dev Container.

In the terminal run:

```bash
cd examples
gitget 'https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/calc/odev_named_range'
```

This will copy the `odev_named_range` example to the examples folder.

In the terminal run:

```bash
cd odev_named_range
python -m start
```

## Example Usage

Starts LibreOffice Calc, runs example and saves output to `tmp/garlic_secrets.ods`

```ps
python .\ex\auto\calc\odev_named_range\start.py -o "tmp/garlic_secrets.ods"
```

[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/
