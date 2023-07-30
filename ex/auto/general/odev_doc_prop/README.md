# Doc Properties

This is a basic example that shows how to write document properties from document path to the command line.

This demo uses [OOO Development Tools](https://python-ooo-dev-tools.readthedocs.io/en/latest/) (OooDev).

OooDev makes this demo possible with just a few lines of code.

See Also: [Examining Office](https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part1/chapter03.html)

See [source code](./start.py)

## Automate

### Cross Platform

From this folder.

```sh
python start.py --doc "algs.odp"
```

### Cross Platform

From this folder.

```sh
python start.py --doc "algs.odp"
```

### Linux/Mac

Run from project root folder.

```sh
python ./ex/auto/general/odev_doc_prop/start.py --doc "ex/auto/general/odev_doc_prop/algs.odp"
```

### Windows

Run from project root folder.

```ps
python .\ex\auto\general\odev_doc_prop\start.py --doc "ex/auto/general/odev_doc_prop/algs.odp"
```

## Live LibreOffice Python

Instructions to run this example in [Live-LibreOffice-Python](https://github.com/Amourspirit/live-libreoffice-python).

Start Live-LibreOffice-Python in a Codespace or in a Dev Container.

In the terminal run:

```bash
cd examples
gitget 'https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/general/odev_doc_prop'
```

This will copy the `odev_doc_prop` example to the examples folder.

In the terminal run:

```bash
cd odev_doc_prop
python -m start -h
```

![Properties Screen Shot](https://user-images.githubusercontent.com/4193389/179302791-d8373bd0-7b72-41a3-86b8-dcbd5bac6feb.png)
