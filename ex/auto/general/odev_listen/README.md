# Office Window Listener

<p align="center">
<img src="https://user-images.githubusercontent.com/4193389/204155527-4e975c63-ea78-4591-a659-d9ddafa8970c.png" width="327" height="218">
</p>

This is a basic example that shows how attach a window listener to office.

This script will stay running until office is closed or `ctl+c` is pressed unless `-t` is passed as a parameter.

When the script starts it will call `minimize()` and `maximize()` several times for the purpose of demonstrating event listening

As (*Write in this case*) window is min, max, activated etc. The events are captured and printed to the screen.

This demo uses [OOO Development Tools] (OooDev).

A `main_loop()` method is called that watches until Office is closed.

[OooDev] makes this demo possible with just a few lines of code.

See Also:

- [OOO Development Tools - Chapter 4. Listening, and Other Techniques](https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part1/chapter04.html)
- [OOO Development Tools - Chapter 25. Monitoring Sheets](https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part4/chapter25.html)
- [Calc Modify Listener Example](../../calc/odev_modify_listener/)
- [Calc Select Listener Example](../../calc/odev_select_listener/)

## Listeners

This example has two classes (`DocWindow`, `DocWindowAdapter`) that show different ways to subscribe to Office Window events.

Functionally these two classes are identical.

Both classes are responsible for starting office.

### Using ODEV TopWindowListener Class

[ODEV] has a `TopWindowListener` class that raise events that can be subscribed to.

[doc_window_adapter.py](./doc_window_adapter.py) (`DocWindowAdapter` class) demonstrates how `TopWindowListener` is implemented.
This means there is no need to create a class that inherits from [XTopWindowListener]

### Implementing XTopWindowListener

[doc_window.py](./doc_window.py) (`DocWindow` class) demonstrates implementing [XTopWindowListener]

## Automate

### Command Line Parameters

- `-a` runs demo using `DocWindowAdapter` else `DocWindow` is used.
- `-t` runs demo and automatically closes office.
- `-h` Displays help on options.

### Dev Container

Run from this folder.

```sh
python start.py
```

### Cross Platform

Run from this folder.

```sh
python start.py
```

### Linus/Mac

From project root folder.

```sh
python ./ex/auto/general/odev_listen/start.py
```

### Windows

From project root folder.

```ps
python .\ex\auto\general\odev_listen\start.py
```

### Example console output

User interactions with window are reflected in console window.

Starts Write as a new document and monitors window activity and auto terminates.

```ps
PS D:\Users\user\Python\python-ooouno-ex> python .\ex\auto\general\odev_listen\start.py
Press 'ctl+c' to exit script early.
Loading Office...
Creating Office document swriter
WL: Opened
Rectangle: (8, 54), 1904 -- 978
WL: Opened
Rectangle: (8, 54), 1904 -- 978
WL: Activated
  Titile bar: Untitled 1 - LibreOffice Writer
WL: Activated
  Titile bar: Untitled 1 - LibreOffice Writer
WL: Minimized
WL: Activated
  Titile bar: Untitled 1 - LibreOffice Writer
WL:  De-activated
WL: Normalized
WL: Minimized
WL: Activated
  Titile bar: Untitled 1 - LibreOffice Writer
WL: Minimized
WL: Activated
  Titile bar: Untitled 1 - LibreOffice Writer
WL:  De-activated
WL: Normalized
WL: Minimized
WL: Activated
  Titile bar: Untitled 1 - LibreOffice Writer
WL: Minimized
WL: Activated
  Titile bar: Untitled 1 - LibreOffice Writer
WL:  De-activated
WL: Normalized
WL: Minimized
WL: Activated
  Titile bar: Untitled 1 - LibreOffice Writer
Closing Office
WL: Activated
  Titile bar: Untitled 1 - LibreOffice Writer
WL: Closed
WL: Closed
Office terminated

Exiting by document close.
```

## Live LibreOffice Python

Instructions to run this example in [Live-LibreOffice-Python](https://github.com/Amourspirit/live-libreoffice-python).

Start Live-LibreOffice-Python in a Codespace or in a Dev Container.

In the terminal run:

```bash
cd examples
gitget 'https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/general/odev_listen'
```

This will copy the `odev_listen` example to the examples folder.

In the terminal run:

```bash
cd odev_listen
python -m start
```

[XTopWindowListener]: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1awt_1_1XTopWindowListener.html

[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/
[OooDev]: https://python-ooo-dev-tools.readthedocs.io/en/latest/
