# Office Window Monitor

This is a basic example that shows how attach a Terminate Monitor to office.

This script will stay running until office is closed or `ctl+c` is pressed.

As (*Calc in this case*) window terminates. The events are captured and printed to the screen.

In this case the `DocMonitor` class that is responsible for starting Office and creating instance of [XTerminateAdapter](https://python-ooo-dev-tools.readthedocs.io/en/latest/src/listeners/x_terminate_adapter.html)
which implements [API XTerminateListener](https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1frame_1_1XTerminateListener.html)
that monitors office terminate.

A `main_loop()` method is called that watches until Office is closed.

This demo uses [OOO Development Tools](https://python-ooo-dev-tools.readthedocs.io/en/latest/) (ODEV).

ODEV makes this demo possible with just a few lines of code.

See Also: [Listening, and Other Techniques](https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part1/chapter04.html)

See [source code](./start.py)

## Automate

```sh
python -m main auto --process "ex/auto/general/odev_monitor/start.py"
```

or for auto shutdown

```sh
python -m main auto --process "ex/auto/general/odev_monitor/start.py True"
```

### Example console output

User interactions with window are reflected in console window.

Starts Calc as a new document and monitors window activity.

```python
PS D:\Users\user\Python\python-ooouno-ex> python -m main auto --process "ex/auto/general/odev_monitor/start.py True"
Press 'ctl+c' to exit script early.
Loading Office...
Creating Office document scalc
Closing Office
TL: Starting Closing
TL: Finished Closing
Office terminated

Exiting by document close.
```
