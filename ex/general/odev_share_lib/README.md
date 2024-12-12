# Shared Library Access

Example of importing a python module from share library using [OOO Development Tools]

By default it is not possible to import a shared library.
In order to solve this the path to the LibreOffice share python directory
must be added to python's sys.path.

[OOO Development Tools] can discover and register the path automatically as shown in this example.

There is a command line example of importing from shared library [here](../../auto/general/odev_share_lib/).

See:

- [Importing Python Modules]
- [Getting Session Information]

## Requirements

For this demo to run it requires a module in "My Modules" named `pyglobal` with the following contents.
This module is in the LibreOffice user's directory of your local machine.

```py
# coding: utf-8
from __future__ import unicode_literals

G_COUNT = 100
```

## Sample Document

see sample LibreOffice Calc document, [share_lib.ods](share_lib.ods)

### Source

see [counter.py](./counter.py)

There is an alternative way to add path. See the [alternative.py](./alternative.py) source code.

## Build

Build will compile the python scripts for this example into a single python script.

The following command will compile script as `share_lib.py` and embed it into`share_lib.ods`
The output is written into `build` folder in the projects root.

```sh
oooscript compile --pyz-out --embed --config "ex/general/odev_share_lib/config.json" --embed-doc "ex/general/odev_share_lib/share_lib.ods"
```

## Demo

Basic demo of a macro that imports a shared python module and also shares it module level `pyglobal.G_COUNT` between documents.

https://user-images.githubusercontent.com/4193389/180564481-b4365aaa-041d-404b-89f0-f9ee66954dcd.mp4

[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/
[Importing Python Modules]: https://help.libreoffice.org/latest/lo/text/sbasic/python/python_import.html
[Getting Session Information]: https://help.libreoffice.org/latest/lo/text/sbasic/python/python_session.html
