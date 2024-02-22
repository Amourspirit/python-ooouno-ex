# Calc Convert link text to Hyperlink

<p align="center">
<img src="https://github.com/Amourspirit/python-ooouno-ex/assets/4193389/2261b2b9-ea30-4068-b37e-e3b846e6759b" width="489" height="269">
</p>


This is a basic example that opens up a new Calc document and inserts array of links as text and then converts them to hyperlinks.

This demo uses [OOO Development Tools](https://python-ooo-dev-tools.readthedocs.io/en/latest/) (OooDev).

See [source code](./start.py)

There is a macro demo of this that can be found in the `macro` folder, which is in the root directory of this repository.
The macro file is called `ooodev_calc_hyperlinks.py`.
To run the macro demo open this project in Codespace or Dev Container and run the following macro in Calc:

`Tools` -> `Macros` -> `Run Macro...` -> `My Macros` -> `ooodev_calc_hyperlinks` -> `make_hyperlinks`

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
python ./ex/auto/calc/odev_hyperlinks/start.py
```

### Windows

```ps
python .\ex\auto\calc\odev_hyperlinks\start.py
```

## Live LibreOffice Python

Instructions to run this example in [Live-LibreOffice-Python](https://github.com/Amourspirit/live-libreoffice-python).

Start Live-LibreOffice-Python in a Codespace or in a Dev Container.

In the terminal run:

```bash
cd examples
gitget 'https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/calc/odev_hyperlinks'
```

This will copy the `odev_hyperlinks` example to the examples folder.

In the terminal run:

```bash
cd odev_hyperlinks
python -m start
```
