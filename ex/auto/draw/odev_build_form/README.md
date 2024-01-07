<p align="center">
<img src="https://github.com/Amourspirit/python-ooouno-ex/assets/4193389/cf9ef0e0-33dd-4210-8310-d0c2bbe00edc" width="600" height="573" alt="form"/>
</p>

# Build Form

Demonstrates how to build a form from code and attach form to a database.

This example uses [OOO Development Tools] (OooDev)

See [start source code](./start.py) and [build_form source code](./build_form.py)

This example uses the form controls in the `ooodev.form.controls` modules by using the various `ooodev.Draw.DrawForm` insert control methods. By using the `ooodev.form.controls` modules you can create controls without having to implement listeners on the class. The controls can subscribe to any event that it supports.

This example also demonstrates adding a rectangle to the draw page as a background to the form.

See also [Build Form2](../../writer/odev_build_form2/) that demonstrates how to build a form from code and attach form to a database in Writer.

**Example code for creating a form and inserting controls**

```python
def create_form(self) -> None:
    # ...
    draw_page = self._doc.slides[0]
    # ...
    main_form = draw_page.forms.add_form("MainForm")
    # ...
    self._ctl_lbl_age = main_form.insert_control_label(
        label="AGE",
        x=x1,
        y=y,
        width=width1,
        height=height,
        styles=[font_colored],
    )

    self._ctl_num_age = main_form.insert_control_numeric_field(
        x=x2,
        y=y,
        width=width2,
        height=height,
        accuracy=0,
        spin_button=False,
        border=BorderKind.BORDER_3D,
    )
    # listen for events
    self._ctl_lbl_age.bind_to_control(self._ctl_num_age)
    self._ctl_num_age.add_event_down(self._fn_on_down)
    self._ctl_num_age.add_event_up(self._fn_on_up)
    self._set_tab_index(self._ctl_num_age)
    # ...
```

## Automate

This example when run will wait for you to close the document and will print various event information to the console.

### Dev Container

From current example folder.

```shell
python -m start
```

### Cross Platform

From current example folder.

```shell
python -m start
```

### Linux/Mac

From project root folder

```sh
python ./ex/auto/draw/odev_build_form/start.py
```

### Windows

From project root folder

```ps
python .\ex\auto\draw\odev_build_form\start.py
```

## Live LibreOffice Python

Instructions to run this example in [Live-LibreOffice-Python](https://github.com/Amourspirit/live-libreoffice-python).

Start Live-LibreOffice-Python in a Codespace or in a Dev Container.

In the terminal run:

```bash
cd examples
gitget 'https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/draw/odev_build_form'
```

This will copy the `odev_build_form` example to the examples folder.

In the terminal run:

```bash
cd odev_build_form
python -m start
```

## Note

If you get an error `No SDBC driver was found for the URL 'sdbc:embedded:hsqldb'.` you most likely need to enable Java in LibreOffice.

[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/