# Calc Sheet Activation Event Example

<p align="center">
<img src="https://github.com/Amourspirit/python-ooouno-ex/assets/4193389/65e447b1-5493-479e-9a2a-5ad7f4f15d14" width="600" height="399">
</p>

Simple example of being notified of Spreadsheet Activation changes.

When the spreadsheet is activated the sheet name is outputted to console and a Message Box is displayed.

This script will stay running until office is closed or `ctl+c` is pressed unless `-t` is passed as a parameter.

This demo uses [OOO Development Tools] (OooDev).

A `main_loop()` method is called that watches until Office is closed.

## Automate

### Command Line Parameters

- `-t` runs demo and automatically closes office.
- `-h` Displays help on options.

### Dev Container

Run from this folder.

```sh
python -m start
```

### Cross Platform

Run from this folder.

```sh
python -m start
```

### Linus/Mac

From project root folder.

```sh
python ./ex/auto/calc/odev_sheet_activation_event/start.py
```

### Windows

From project root folder.

```ps
python .\ex\auto\calc\odev_sheet_activation_event\start.py
```

## Live LibreOffice Python

Instructions to run this example in [Live-LibreOffice-Python](https://github.com/Amourspirit/live-libreoffice-python).

Start Live-LibreOffice-Python in a Codespace or in a Dev Container.

In the terminal run:

```bash
cd examples
gitget 'https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/calc/odev_sheet_activation_event'
```

This will copy the `odev_sheet_activation_event` example to the examples folder.

In the terminal run:

```bash
cd odev_sheet_activation_event
python -m start
```

[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/
[OooDev]: https://python-ooo-dev-tools.readthedocs.io/en/latest/
