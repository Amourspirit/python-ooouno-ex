# Doc Info

This is a basic example that shows how to write document information from document path to the command line.

This demo uses [OOO Development Tools](https://python-ooo-dev-tools.readthedocs.io/en/latest/) (OooDev).

See Also: [Examining Office](https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part1/chapter03.html)

See [source code](./start.py)

## Automate

### Dev Container

From project root folder.

```sh
python start.py --doc "story.odt" --service --interface --xdoc --property
```

### Cross Platform

From project root folder.

```sh
python start.py --doc "story.odt" --service --interface --xdoc --property
```

or, to see all possible commands

```shell
python start.py -h
```

### Linux

From project root folder.

```sh
python ./ex/auto/general/odev_doc_info/start.py --doc "ex/auto/general/odev_doc_info/story.odt" --service --interface --xdoc --property
```

### Windows

```ps
python .\ex\auto\general\odev_doc_info\start.py --doc "ex\auto\general\odev_doc_info\story.odt" --service --interface --xdoc --property
```

## Live LibreOffice Python

Instructions to run this example in [Live-LibreOffice-Python](https://github.com/Amourspirit/live-libreoffice-python).

Start Live-LibreOffice-Python in a Codespace or in a Dev Container.

In the terminal run:

```bash
cd examples
gitget 'https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/general/odev_doc_info'
```

This will copy the `odev_doc_info` example to the examples folder.

In the terminal run:

```bash
cd odev_doc_info
python -m start -h
```

![Terminal Example](https://user-images.githubusercontent.com/4193389/179373247-0b9d34b2-9457-44c8-8823-e405272d3c80.gif)
