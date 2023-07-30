# Calc Select Listener

<p align="center">
<img src="https://user-images.githubusercontent.com/4193389/204155527-4e975c63-ea78-4591-a659-d9ddafa8970c.png" width="327" height="218">
</p>

Example of being notified of document changes by attaching to the document's [XModifyBroadcaster].

As cells in spreadsheet are selected the cell value is outputted to console.

Also demos how to attach a window listener to office.

This script will stay running until office is closed or `ctl+c` is pressed unless `-t` is passed as a parameter.

This demo uses [OOO Development Tools] (ODEV).

A `main_loop()` method is called that watches until Office is closed.

See Also:

- [OOO Development Tools - Chapter 25. Monitoring Sheets](https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part4/chapter25.html)
- [OOO Development Tools - Chapter 4. Listening, and Other Techniques](https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part1/chapter04.html)
- [Calc Modify Listener Example](../odev_modify_listener/)
- [Office Window Listener Example](../../general/odev_listen/)

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
python ./ex/auto/calc/odev_select_listener/start.py
```

### Windows

From project root folder.

```ps
python .\ex\auto\calc\odev_select_listener\start.py
```

### Example console output

User interactions with window are reflected in console window.

Starts Write as a new document and monitors window activity and auto terminates.

```ps
PS D:\Users\user\Python\python-ooouno-ex> python .\ex\auto\calc\odev_select_listener\start.py
Press 'ctl+c' to exit script early.
Loading Office...
Creating Office document scalc
A2 value: 42.0
A3 value: 58.9
A4 value: -66.5
A5 value: 43.4
A6 value: 44.5
A7 value: 45.3
Closing
Closing the document
A7 value: 45.3
A7 value: 45.3
Closing Office
Office terminated

Exiting by document close.
```

## Live LibreOffice Python

Instructions to run this example in [Live-LibreOffice-Python](https://github.com/Amourspirit/live-libreoffice-python).

Start Live-LibreOffice-Python in a Codespace or in a Dev Container.

In the terminal run:

```bash
cd examples
gitget 'https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/calc/odev_select_listener'
```

This will copy the `odev_select_listener` example to the examples folder.

In the terminal run:

```bash
cd odev_select_listener
python -m start
```

[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/
[ODEV]: https://python-ooo-dev-tools.readthedocs.io/en/latest/

[XModifyBroadcaster]: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1util_1_1XModifyBroadcaster.html