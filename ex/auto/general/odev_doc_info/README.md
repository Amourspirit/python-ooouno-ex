# Doc Info

This is a basic example that shows how to write document information from document path to the command line.

This demo uses [OOO Development Tools](https://python-ooo-dev-tools.readthedocs.io/en/latest/) (ODEV).

See Also: [Examining Office](https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part1/chapter03.html)

See [source code](./start.py)

## Automate

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

![Terminal Example](https://user-images.githubusercontent.com/4193389/179373247-0b9d34b2-9457-44c8-8823-e405272d3c80.gif)
