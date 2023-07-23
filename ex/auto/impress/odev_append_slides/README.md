# Impress append Slides to existing slide show

<p align="center">
  <img width="435" height="448" src="https://user-images.githubusercontent.com/4193389/198401485-94062f29-6a24-40f7-8873-fce8abaff481.png">
</p>

This example demonstrates how to combine Slide show documents using Impress.

One challenge is for every slide that is appended a confirmation dialog pops up needing to click **yes** ( five time when adding `points.odp` ) in a dialog prompt.
This is where [ooo-dev-tools-gui-win] come in, it can monitor for a dialog box and press a button automatically.

You can see here how [ooo-dev-tools-gui-win] is called.

```python
try:
    # only in windows
    from odevgui_win.dialog_auto import DialogAuto
except ImportError:
    DialogAuto = None

# partial AppendSlides class
class AppendSlides:
    # ...

    def append(self) -> None:
      # other code
      # ...

      # monitor for Confirmation dialog
      if DialogAuto:
          DialogAuto.monitor_dialog('y')

      # rest of the code ...
```

There is one limitation at this time.
[ooo-dev-tools-gui-win] is only for [OOO Development Tools] on windows.

This demo uses [OOO Development Tools]

See Also:

- [OOO Development Tools - Chapter 17. Slide Deck Manipulation](https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part3/chapter17.html)

## Automate

An extra parameters can be passed in:

The first parameter would be the slide show file to append to.

All successive files are append to the first.

**Example:**

```sh
python ./ex/auto/impress/odev_append_slides/start.py "resources/presentation/algs.odp" "resources/presentation/points.odp"
```

If no args are passed in then the `points.odp` is appended to `algs.odp`.

The document is not saved by default.

A message box is display once the document has been created asking if you want to close the document.

### Cross Platform

From current example folder.

```sh
python -m start
```

### Linux/Mac

```sh
python ./ex/auto/impress/odev_append_slides/start.py
```

### Windows

```ps
python .\ex\auto\impress\odev_append_slides\start.py
```

## Live LibreOffice Python

Instructions to run this example in [Live-LibreOffice-Python](https://github.com/Amourspirit/live-libreoffice-python).

Start Live-LibreOffice-Python in a Codespace or in a Dev Container.

In the terminal run:

```bash
cd examples
gitget 'https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/impress/odev_append_slides'
```

This will copy the `odev_append_slides` example to the examples folder.

In the terminal run:

```bash
cd odev_append_slides
python -m start
```

[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/
[ooo-dev-tools-gui-win]: https://ooo-dev-tools-gui-win.readthedocs.io/en/latest/index.html
