# Export Calc Sheet Range as Image

<p align="center">
<img src="https://github.com/Amourspirit/python-ooouno-ex/assets/4193389/5cdf724c-1ba4-42ac-a14a-4a2de2d1318d">
</p>

Demonstrates saving a Calc sheet range as an image.

This demo uses This demo uses [OOO Development Tools] (OooDev).

It only takes a few lines of code to export a Calc sheet range as an image.

```python
rng = sheet.get_range(range_name="A1:N4")
rng.export_as_image(file)
```

if needed you can modify the image filters via the events.

```python
from ooodev.events.args.cancel_event_args_generic import CancelEventArgsGeneric
from ooodev.events.event_data.img_export_t import ImgExportT
from ooodev.calc import CalcNamedEvent

def on_exporting(source: Any, args: CancelEventArgsGeneric[ImgExportT]) -> None:
    # set the image compression to 9
    args.event_data["compression"] = 9

rng = sheet.get_range(range_name="A1:N4")
rng.subscribe_event(CalcNamedEvent.RANGE_EXPORTING_IMAGE, on_exporting)

# on_exporting will be called when the image is exported
rng.export_as_image(file)
```

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
