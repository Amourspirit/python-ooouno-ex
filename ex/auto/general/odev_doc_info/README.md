# Doc Info

This is a basic example that shows how to write document information from document path to the command line.

This demo uses [OOO Development Tools](https://python-ooo-dev-tools.readthedocs.io/en/latest/) (ODEV).

See Also: [Examining Office](https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part1/chapter03.html)

See [source code](./start.py)

## Automate

```sh
python -m main auto --process "ex/auto/general/odev_doc_info/start.py --doc ex/auto/general/odev_doc_info/story.odt --service --interface --xdoc --property"
```

or, to see all possible commands

```sh
python -m main auto --process "ex/auto/general/odev_doc_info/start.py -h"
```

![Terminal Example](https://user-images.githubusercontent.com/4193389/179373247-0b9d34b2-9457-44c8-8823-e405272d3c80.gif)
