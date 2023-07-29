# Set page margins

In this example we will set the page margins of a document.

This example takes advantage of the hundreds of format styles that [OOO Development Tools] has to offer which makes setting the page margins is a simple task.

See [source code](./start.py)

## Automate

### Dev Container

From this folder.

```sh
python -m start --file "./data/cicero_dummy.odt"
```

### Cross Platform

From this folder.

```sh
python -m start --file "./data/cicero_dummy.odt"
```

### Linux/Mac

Run from current example folder.

From project root folder (for default file).

```sh
python ./ex/auto/writer/odev_formatting/page_margin/start.py
```

### Windows

From project root folder (for default file).

```ps
python .\ex\auto\writer\odev_formatting\page_margin\start.py
```

## Live LibreOffice Python

Instructions to run this example in [Live-LibreOffice-Python](https://github.com/Amourspirit/live-libreoffice-python).

Start Live-LibreOffice-Python in a Codespace or in a Dev Container.

In the terminal run:

```bash
cd examples
gitget 'https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/writer/odev_formatting/page_margin'
```

This will copy the `page_margin` example to the examples folder.

In the terminal run:

```bash
cd page_margin
python -m start
```

## Screenshots

### Before

![2023-04-22_12-16-06](https://user-images.githubusercontent.com/4193389/233795486-fba0bbd6-cb9f-4ae5-a475-10c6cfb22feb.png)

### After

![2023-04-22_12-17-28](https://user-images.githubusercontent.com/4193389/233795531-20dacbf7-3a20-4295-8308-0635deff6668.png)


[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/
