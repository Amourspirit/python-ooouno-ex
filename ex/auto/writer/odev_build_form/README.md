# Build Form

Demonstrates how to build a form from code and attach form to a database.

This example uses [OOO Development Tools]

See [start source code](./start.py) and [build_form source code](./build_form.py)

There is a [Build Form2](../odev_build_form2) example that uses the `ooodev.form.controls` modules to create controls.
The `Form2` example is the recommended way to create forms with controls.

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
python ./ex/auto/writer/odev_build_form/start.py
```

### Windows

From project root folder

```ps
python .\ex\auto\writer\odev_build_form\start.py
```

![Form-screenshot](https://user-images.githubusercontent.com/4193389/194674585-8252bf4b-3ada-4746-a70a-234e91767b85.png)

## Live LibreOffice Python

Instructions to run this example in [Live-LibreOffice-Python](https://github.com/Amourspirit/live-libreoffice-python).

Start Live-LibreOffice-Python in a Codespace or in a Dev Container.

In the terminal run:

```bash
cd examples
gitget 'https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/writer/odev_build_form'
```

This will copy the `odev_build_form` example to the examples folder.

In the terminal run:

```bash
cd odev_build_form
python -m start
```

[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/