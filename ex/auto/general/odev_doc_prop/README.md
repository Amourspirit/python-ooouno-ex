# Doc Properties

This is a basic example that shows how to write document properties from document path to the command line.

This demo uses [OOO Development Tools](https://python-ooo-dev-tools.readthedocs.io/en/latest/) (ODEV).

ODEV makes this demo possible with just a few lines of code.

See Also: [Examining Office](https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part1/chapter03.html)

See [source code](./start.py)

## Automate

### Cross Platform

From project root folder.

```shell
python -m main auto -p "ex/auto/general/odev_doc_prop/start.py --doc ex/auto/general/odev_doc_prop/algs.odp"
```

### Linux

Run from current example folder.

```shell
python start.py --doc "algs.odp"
```

![Properties Screen Shot](https://user-images.githubusercontent.com/4193389/179302791-d8373bd0-7b72-41a3-86b8-dcbd5bac6feb.png)
