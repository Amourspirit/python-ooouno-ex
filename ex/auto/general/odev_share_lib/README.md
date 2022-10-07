# Shared Library Access

Example of importing a python module from share library using [OOO Development Tools] (ODEV)

The path of LibreOffice user shared python Libraries can be different from
user to user and machine to machine.

Also ODEV has options to set these paths to a new location temporarily.

Getting the locations of these paths can be challenging.

ODEV can discover and register the path automatically as shown in this example.

Note that ODEV requires a connection to LibreOffice as it is the LibreOffice API
that allows then paths to be discovered. Without a connection automatic registering of path is not possible.

There is a macro example of importing from shared library [here](../../../general/odev_share_lib/).

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

see [start.py](./start.py)

There is an alternative way to add path. See the [alternative.py](./alternative.py) source code.

## Automate

### Cross Platform

From this folder.

```sh
python -m start
```

### Linus/Mac

From project root folder

```sh
python ./ex/auto/general/odev_share_lib/start.py
```

### Windows

From project root folder

```ps
python .\ex\auto\general\odev_share_lib\start.py
```

### Example console output

```text
PS D:\Users\user\Python\python-ooouno-ex> python -m main auto --process "ex/auto/general/odev_share_lib/start.py"
As expected unable to import pyglobal without registering path.

As expected unable to register path before Lo.load_office is called

Loading Office...
Closing Office
Office terminated
100
101
102
103
104
105
106
107
108
109
110
```

[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/
[Importing Python Modules]: https://help.libreoffice.org/latest/lo/text/sbasic/python/python_import.html
[Getting Session Information]: https://help.libreoffice.org/latest/lo/text/sbasic/python/python_session.html
