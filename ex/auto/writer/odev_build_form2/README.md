<p align="center">
<img src="https://user-images.githubusercontent.com/4193389/194674585-8252bf4b-3ada-4746-a70a-234e91767b85.png" width="558" height="552" alt="form"/>
</p>

# Build Form2

Demonstrates how to build a form from code and attach form to a database.

This example uses [OOO Development Tools] (OooDev)

See [start source code](./start.py) and [build_form source code](./build_form.py)

Unlike [Build Form](../odev_build_form) this example uses the form controls in the `ooodev.form.controls` modules by using the
various `ooodev.utils.forms.Forms` insert control methods. By using the `ooodev.form.controls` modules you can create
controls without having to implement listeners on the class. The controls can subscribe to any event that it supports.

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
python ./ex/auto/writer/odev_build_form2/start.py
```

### Windows

From project root folder

```ps
python .\ex\auto\writer\odev_build_form2\start.py
```

## Live LibreOffice Python

Instructions to run this example in [Live-LibreOffice-Python](https://github.com/Amourspirit/live-libreoffice-python).

Start Live-LibreOffice-Python in a Codespace or in a Dev Container.

In the terminal run:

```bash
cd examples
gitget 'https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/writer/odev_build_form2'
```

This will copy the `odev_build_form2` example to the examples folder.

In the terminal run:

```bash
cd odev_build_form2
python -m start
```

## Note

If you get an error `No SDBC driver was found for the URL 'sdbc:embedded:hsqldb'.` you most likely need to enable Java in LibreOffice.


[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/