# Office Window Listener

This is a basic example that shows how attach a window listener to office.

This script will stay running until office is closed or `ctl+c` is pressed.

As (*Write in this case*) window is min, max, activated etc. The events are captured and printed to the screen.

In this case the `DocWindow` class that is responsible for starting Office is implements [XTopWindowAdapter](https://python-ooo-dev-tools.readthedocs.io/en/latest/src/listeners/x_top_window_adapter.html)
which implements [API XTopWindowListener](https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1awt_1_1XTopWindowListener.html).

A `main_loop()` method is called that watches until Office is closed.

This demo uses [OOO Development Tools](https://python-ooo-dev-tools.readthedocs.io/en/latest/) (ODEV).

ODEV makes this demo possible with just a few lines of code.

See Also: [Listening, and Other Techniques](https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part1/chapter04.html)

See [source code](./start.py)

## Automate

```sh
python -m main auto --process "ex/auto/general/odev_listen/start.py"
```

### Example console output

User interactions with window are reflected in console window.

Starts Write as a new document and monitors window activity.

```python
PS D:\Users\Python\python-ooouno-ex> python -m main auto --process "ex/auto/general/odev_listen/start.py"
Loading Office...
Creating Office document swriter
WL: Opened
Rectangle: (-2, 58), 1904 -- 978
WL: Opened
Rectangle: (-2, 58), 1904 -- 978
WL: Activated
  Titile bar: Untitled 1 - LibreOffice Writer
WL: Activated
  Titile bar: Untitled 1 - LibreOffice Writer
WL:  De-activated
WL: Minimized
WL: Normalized
WL: Activated
  Titile bar: Untitled 1 - LibreOffice Writer
WL:  De-activated
WL: Minimized
WL: Normalized
WL: Activated
  Titile bar: Untitled 1 - LibreOffice Writer
WL: Closing
WL: Activated
  Titile bar: Untitled 1 - LibreOffice Writer
WL: Closed
WL: Closed

Exiting by document close.
```
