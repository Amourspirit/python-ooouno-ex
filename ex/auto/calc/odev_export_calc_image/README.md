# Export Calc Sheet Range as Image

<p align="center">
<img src="https://github.com/Amourspirit/python-ooouno-ex/assets/4193389/5cdf724c-1ba4-42ac-a14a-4a2de2d1318d">
</p>

## Overview

Demonstrates saving a Calc sheet range as an image.

This demo uses This demo uses [OOO Development Tools] (OooDev).

It only takes a few lines of code to export a Calc sheet range as an image.

```python
rng = sheet.get_range(range_name="A1:N4")
rng.export_as_image("my_image.png")
```

if needed you can modify the image filters via the events.

```python
from typing import TYPE_CHECKING
if TYPE_CHECKING:
    # Any imports in this block are only needed for type checking at design time and
    # are not available at runtime.
    from ooodev.events.args.cancel_event_args_export import CancelEventArgsExport
    from ooodev.events.args.event_args_export import EventArgsExport
    from ooodev.calc.filter.export_jpg import ExportJpgT
    from ooodev.calc.filter.export_png import ExportPngT


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



rng = sheet.get_range(range_name="A1:N4")
# Register event handlers so we can have a little more fine control over the export.
# It is not required to register event handlers to export an image.
rng.subscribe_event(CalcNamedEvent.EXPORTING_RANGE_JPG, on_exporting_jpg)
rng.subscribe_event(CalcNamedEvent.EXPORTED_RANGE_JPG, on_exported_jpg)
rng.subscribe_event(CalcNamedEvent.EXPORTING_RANGE_PNG, on_exporting_png)
rng.subscribe_event(CalcNamedEvent.EXPORTED_RANGE_PNG, on_exported_png)

# on_exporting_png() will be called when the image is exported
rng.export_as_image(fnm="my_image.png", resolution=200)

# on_exporting_jpg() will be called when the image is exported as a jpg file.
```

### See Also

- [export_as_image()](https://python-ooo-dev-tools.readthedocs.io/en/latest/src/calc/calc_cell_range.html#ooodev.calc.CalcCellRange.export_as_image)
- [export_png()](https://python-ooo-dev-tools.readthedocs.io/en/latest/src/calc/calc_cell_range.html#ooodev.calc.CalcCellRange.export_png)
- [export_jpg()](https://python-ooo-dev-tools.readthedocs.io/en/latest/src/calc/calc_cell_range.html#ooodev.calc.CalcCellRange.export_jpg)

## Automate

A message box is display once the document has been processed asking if you want to close the document.

### Dev Container

From this folder.

```sh
python -m start
```

### Cross Platform

From this folder.

```sh
python -m start
```

### Linux/Mac

```sh
python ./ex/auto/calc/odev_export_calc_image/start.py
```


### Windows

```ps
python .\ex\auto\calc\odev_export_calc_image\start.py
```

## Live LibreOffice Python

Instructions to run this example in [Live-LibreOffice-Python](https://github.com/Amourspirit/live-libreoffice-python).

Start Live-LibreOffice-Python in a Codespace or in a Dev Container.

In the terminal run:

```bash
cd examples
gitget 'https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/calc/odev_export_calc_image'
```

This will copy the `odev_export_calc_image` example to the examples folder.

In the terminal run:

```bash
cd odev_export_calc_image
python -m start
```

[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/
