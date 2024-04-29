<p align="center">
<img src="https://github.com/Amourspirit/python-ooouno-ex/assets/4193389/832195c4-9e7c-403a-9600-0dbe52cf0405" width="550" height="488">
</p>

# Calc Add Range of Data Automation

This is a basic example that opens up a new Calc document and inserts array of data and a formula and freezes the first row of the sheet.

This demo uses [OOO Development Tools](https://python-ooo-dev-tools.readthedocs.io/en/latest/) (OooDev).

OooDev makes this demo possible with just a few lines of code.

Also demonstrates saving an document, displaying a message box and displaying a info bar message.

See [source code](./start.py)

There is a macro demo of this. See [macro example](../../../calc/odev_add_range_data)

## Automate

A message box is display once the document has been created asking if you want to close the document.

The following command will run automation that generates a new Calc with data.

### Dev Container

From this folder.

```sh
python -m start
```

### Cross platform

From this folder.

```sh
python -m start
```

### Linux/Mac

```sh
python ./ex/auto/calc/odev_add_range_data/start.py
```

### Windows

```ps
python .\ex\auto\calc\odev_add_range_data\start.py
```

## Live LibreOffice Python

Instructions to run this example in [Live-LibreOffice-Python](https://github.com/Amourspirit/live-libreoffice-python).

Start Live-LibreOffice-Python in a Codespace or in a Dev Container.

In the terminal run:

```bash
cd examples
gitget 'https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/calc/odev_add_range_data'
```

This will copy the `odev_add_range_data` example to the examples folder.

In the terminal run:

```bash
cd odev_add_range_data
python -m start
```

