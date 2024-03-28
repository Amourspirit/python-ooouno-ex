# Hello World Automation

This is a basic example that opens up a new Writer document and writes *Hello World* using [ScriptForge](https://gitlab.com/LibreOfficiant/scriptforge)
and [ScriptForge Typings](https://pypi.org/project/types-scriptforge/).
ScriptForge is a repository of "macro scripting resources" that are written in basic, but callable from python. They have been contributed for consideration to be incorporated in future distributions of LibreOffice (hence the "forge" aspect of the name.) In the mean time, we can access them explicitly via the scriptforge PyPI package. See https://gitlab.com/LibreOfficiant/scriptforge.

NOTE: This is not a real-world example of why you would want to invoke any of the ScriptForge resources. It merely illustrates how you would go about it should you decide that it is useful to you.

See [source code](./start.py)

## Automate


The following command will run automation that generates a new Writer document and writes "Hello World" into it.

### Dev Container

Run from this example folder.

```sh
python -m start
```

### Cross Platform

Run from this example folder.

```sh
python -m start
```

### Linux/Mac

From project root folder.

```sh
python ./ex/auto/writer/hello_world/start.py
```

### Windows

From project root folder.

```ps
python .\ex\auto\writer\hello_world\start.py
```

## Live LibreOffice Python

Instructions to run this example in [Live-LibreOffice-Python](https://github.com/Amourspirit/live-libreoffice-python).

Start Live-LibreOffice-Python in a Codespace or in a Dev Container.

In the terminal run:

```bash
cd examples
gitget 'https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/writer/hello_world'
```

This will copy the `hello_world` example to the examples folder.

In the terminal run:

```bash
cd hello_world
python -m start
```
