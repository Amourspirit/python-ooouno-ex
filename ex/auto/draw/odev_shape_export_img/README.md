# Export Shape as Image

<p align="center">
<img src="https://github.com/Amourspirit/python-ooouno-ex/assets/4193389/3de0b934-ad2c-403c-8b13-dd24552da5cd", width="330" height="325">
</p>

## Overview

Demonstrates drawing a bezier closed curve in a Draw document and export it as in image.

Building from the [Bezier Closed example](../odev_bezier_closed/) this example adds the ability to export the shape as an image.

A message box is display once the document has been created asking if you want to close the document.

This demo uses [OOO Development Tools]

All shapes in the [ooodev.draw.shapes](https://python-ooo-dev-tools.readthedocs.io/en/latest/src/draw/shapes/index.html) have `export_shape_png()` and `export_shape_jpg()` methods.

See [OOO Development Tools: Chapter 15. Complex Shapes](https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part3/chapter15.html)

If needed you can modify the image filters via the events.

```python
from typing import TYPE_CHECKING
if TYPE_CHECKING:
    # Any imports in this block are only needed for type checking at design time and
    # are not available at runtime.
    from ooodev.events.args.cancel_event_args_export import CancelEventArgsExport
    from ooodev.events.args.event_args_export import EventArgsExport
    from ooodev.draw.filter.export_jpg import ExportJpgT
    from ooodev.draw.filter.export_png import ExportPngT


# region event handlers
def on_exporting_png(source: Any, args: CancelEventArgsExport[ExportPngT]) -> None:
    args.event_data["translucent"] = False
    args.event_data["compression"] = 8  # 0-9

def on_exported_png(source: Any, args: EventArgsExport[ExportPngT]) -> None:
    print(f'Png URL: {args.get("url")}')

def on_exporting_jpg(source: Any, args: CancelEventArgsExport[ExportJpgT]) -> None:
    args.event_data["quality"] = 80  # 0-100
    # when color_mode False image is exported as grayscale.
    args.event_data["color_mode"] = False

def on_exported_jpg(source: Any, args: EventArgsExport[ExportJpgT]) -> None:
    print(f'Jpg URL: {args.get("url")}')
# endregion event handlers

shape = self._create_bezier(slide=slide)
# Register event handlers so we can have a little more fine control over the export.
# It is not required to register event handlers to export an image.
shape.subscribe_event_shape_jpg_exporting(on_exporting_png)
shape.subscribe_event_shape_png_exported(on_exported_png)
shape.subscribe_event_shape_png_exporting(on_exporting_jpg)
shape.subscribe_event_shape_jpg_exported(on_exported_jpg)

# on_exporting_png() will be called when the image is exported
if ext == "png":
    shape.export_shape_png(self._fnm, resolution=resolution)
else:
    shape.export_shape_jpg(self._fnm, resolution=resolution)

# on_exporting_jpg() will be called when the image is exported as a jpg file.
```

### See Also

- [ClosedBezierShape.export_shape_png()](https://python-ooo-dev-tools.readthedocs.io/en/latest/src/draw/shapes/closed_bezier_shape.html#ooodev.draw.shapes.ClosedBezierShape.export_shape_png)
- [ClosedBezierShape.export_shape_jpg()](https://python-ooo-dev-tools.readthedocs.io/en/latest/src/draw/shapes/closed_bezier_shape.html#ooodev.draw.shapes.ClosedBezierShape.export_shape_jpg)

## Automate

An extra parameter can be passed in:

A value of `jpg` or `png` The default value is `png`.

- `png` - Export the shape as a png image.
- `jpg` - Export the shape as a jpg image.

**Example:**

```sh
python -m start jpg
```

### Dev Container

From current example folder.

```sh
python -m start
```

### Cross Platform

From current example folder.

```sh
python -m start
```

### Linux/Mac

```sh
python ./ex/auto/draw/odev_shape_export_img/start.py
```

### Windows

```ps
python .\ex\auto\draw\odev_shape_export_img\start.py
```

[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/

## Live LibreOffice Python

Instructions to run this example in [Live-LibreOffice-Python](https://github.com/Amourspirit/live-libreoffice-python).

Start Live-LibreOffice-Python in a Codespace or in a Dev Container.

In the terminal run:

```bash
cd examples
gitget 'https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/draw/odev_shape_export_img'
```

This will copy the `odev_shape_export_img` example to the examples folder.

In the terminal run:

```bash
cd odev_shape_export_img
python -m start
```
