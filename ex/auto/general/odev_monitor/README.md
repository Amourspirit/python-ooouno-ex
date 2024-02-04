# Office Window Monitor

<p align="center">
<img src="https://user-images.githubusercontent.com/4193389/204155527-4e975c63-ea78-4591-a659-d9ddafa8970c.png" width="327" height="218">
</p>

This is a basic example that shows how attach a Terminate Monitor to office.
In addition a listener is attached to bridge connection to office and
if the bridge terminates for any reason before office is closed then script will also terminate.

This script will stay running until office is closed or `ctl+c` is pressed. Note that closing the window does not trigger termination.
Must use `File --> Exit LibreOffice` or `ctl+q` to close office in order to trigger termination.

As (*Calc in this case*) window terminates. The events are captured and printed to the screen.

In this case the `DocMonitor` class that is responsible for starting Office and creating instance of `TerminateListener` which implements [API XTerminateListener](https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1frame_1_1XTerminateListener.html)
that monitors office terminate.

`DocMonitor` also has two additional ways of listening if the bridge to LibreOffice terminates for some
unexpected reason. One way is to listen for an event raised by `Lo` when bridge is gone.
Another way is to attach a listener to `Lo.bridge`. Only one of these methods is necessary to be notified
when bridge is gone away but both are included for example purposes.

A `main_loop()` method is called that watches until Office is closed.

This demo uses [OOO Development Tools](https://python-ooo-dev-tools.readthedocs.io/en/latest/) (OooDev).

See Also:

- [OOO Development Tools - Chapter 4. Listening, and Other Techniques](https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part1/chapter04.html)
- [OOO Development Tools - Chapter 25. Monitoring Sheets](https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part4/chapter25.html)
- [Calc Modify Listener Example](../../calc/odev_modify_listener/)
- [Calc Select Listener Example](../../calc/odev_select_listener/)
- [Office Window Listener Example](../odev_listen/)

## Automate

### Dev Container

Run from this example folder.

```sh
python start.py
```

### Cross Platform

Run from this example folder.

```sh
python start.py
```

or for auto shutdown

```sh
python start.py True
```

### Linux/Mac

From project root folder.

```sh
python ./ex/auto/general/odev_monitor/start.py
```
or for auto shutdown

```sh
python ./ex/auto/general/odev_monitor/start.py True
```

### Windows

From project root folder.

```ps
python .\ex\auto\general\odev_monitor\start.py
```

or for auto shutdown

```ps
python .\ex\auto\general\odev_monitor\start.py True
```

### Example console output

User interactions with window are reflected in console window.

Starts Calc as a new document and monitors window activity.

```text
PS D:\Users\user\Python\python-ooouno-ex> python .\ex\auto\general\odev_monitor\start.py True
Press 'ctl+c' to exit script early.
Loading Office...
Creating Office document scalc
Closing Office
TL: Starting Closing
TL: Finished Closing
Office terminated
Office bridge has gone!!
LO: Office bridge has gone!!
TL: Disposing
```

## Live LibreOffice Python

Instructions to run this example in [Live-LibreOffice-Python](https://github.com/Amourspirit/live-libreoffice-python).

Start Live-LibreOffice-Python in a Codespace or in a Dev Container.

In the terminal run:

```bash
cd examples
gitget 'https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/general/odev_monitor'
```

This will copy the `odev_monitor` example to the examples folder.

In the terminal run:

```bash
cd odev_monitor
python -m start
```
