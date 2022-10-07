# Office Window Listener

This is a basic example that shows how attach a window listener to office.

This script will stay running until office is closed or `ctl+c` is pressed.

When the script starts it will call `minimize()` and `maximize()` several times for the purpose of demonstrating event listening

As (*Write in this case*) window is min, max, activated etc. The events are captured and printed to the screen.

In this case the `DocWindow` class that is responsible for starting Office that implements [XTopWindowAdapter](https://python-ooo-dev-tools.readthedocs.io/en/latest/src/listeners/x_top_window_adapter.html)
which implements [API XTopWindowListener](https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1awt_1_1XTopWindowListener.html).

A `main_loop()` method is called that watches until Office is closed.

This demo uses [OOO Development Tools](https://python-ooo-dev-tools.readthedocs.io/en/latest/) (ODEV).

ODEV makes this demo possible with just a few lines of code.

See Also: [Listening, and Other Techniques](https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part1/chapter04.html)

See [source code](./start.py)

## Automate

### Cross Platform

Run from this folder.

```sh
python start.py
```

or for auto shutdown

```sh
python start.py True
```
### Linus/Mac

From project root folder.

```sh
python ./ex/auto/general/odev_listen/start.py
```

or for auto shutdown

```sh
python ./ex/auto/general/odev_listen/start.py True
```

### Windows

From project root folder.

```sh
python .\ex\auto\general\odev_listen\start.py
```

or for auto shutdown

```sh
python .\ex\auto\general\odev_listen\start.py True
```

### Example console output

User interactions with window are reflected in console window.

Starts Write as a new document and monitors window activity and auto terminates.

```ps
PS D:\Users\user\Python\python-ooouno-ex> python -m main auto --process "ex/auto/general/odev_listen/start.py True"
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
