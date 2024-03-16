# Live LibreOffice Python UNO Examples

Examples for [ooouno], [OOO Development Tools] (OooDev), [ScriptForge] projects and more.

[OooDev] is a powerful python library for LibreOffice and is also available as an [extension](https://extensions.libreoffice.org/en/extensions/show/41700) for LibreOffice

[ooouno] project has made it possible to have a much more
user friendly way of programming for [LibreOffice](https://www.libreoffice.org/). When you use [OooDev] you also get [ooouno] automatically.

The goal of this project is to have ever expanding practical examples for programming
using python in [LibreOffice](https://www.libreoffice.org/) in a truly [OOP](https://en.wikipedia.org/wiki/Object-oriented_programming) way.

At this time the example can be found in the [ex].
As new example are create/submitted the example [ex] folder will be expanded.

This project uses [oooscript] to compile multiple scripts into a single script and embed output into a document.
This makes it easy to use as a LibreOffice macro.

Another fantastic resource is [LibreOffice Programming],
great documentation and many java examples.
An archive of the Java code is available at https://github.com/Amourspirit/libreoffice_lop_java.

## I NEED A BREAK ...

Checkout the LibreOffice Calc [Sudoku example](./ex/calc/sudoku).

![calc_sudoku](https://user-images.githubusercontent.com/4193389/165391098-883a7647-5fc8-47de-b028-4c2c98337abe.png)

## Running Project Examples

This project is set up to work with [Vs Code](https://code.visualstudio.com/) and [Github Codespaces](https://docs.github.com/en/codespaces/overview) or in a in a Development Container on your local computer.

All of the examples can be run from a web browser using [Github Codespaces](https://docs.github.com/en/codespaces/overview) or in a in a Development Container on your local computer.

This project's Development container is based off [Live LibreOffice Python](https://github.com/Amourspirit/live-libreoffice-python) template. See the [Live LibreOffice Python Wiki](https://github.com/Amourspirit/live-libreoffice-python/wiki) for more information.

## Project Installation

Installing this project is not necessary if you are using [Github Codespaces](https://docs.github.com/en/codespaces/overview) or in a in a Development Container on your local computer, However; if you want to install locally follow the instructions below.

This project use a virtual environment for development purposes.

[poetry] is required to install this project.

In order to run test it is essential that LibreOffice's `uno.py` and `unohelper.py` can be imported. These files are part of the LibreOffice installation. The location of these files vary depending on OS and other factors.

### Linux

Set up virtual environment if not existing.

```sh
python3 -m venv ./.venv
```

Activate virtual environment and install development requirements.

```sh
. ./.venv/bin/activate
```

Install requirement using poetry.

```sh
poetry install
```

Add the `uno.py` and `unohelper.py` links to virtual environment.

```sh
python -m main cmd-link --add
```

### Windows

Windows is a little trickery. Creating a link to `uno.py` and importing it will not work as it does in Linux. This is due to the how LibreOffice implements the python environment on Windows.

There are build in tools that aid in this that we will get to shortly.

The work around on Windows is a slight hack to the virtual environment.

Set up virtual environment if not existing (recommend using PowerShell).

```ps
python -m venv .\.venv
```

Activate virtual environment and install development requirements.

```ps
.\.venv\Scripts\activate
```

Install requirements using poetry.

```ps
poetry install
```

After installing using the previous command it time to set the environment to work with LibreOffice.

```ps
python -m main env -t
```

This will set the virtual environment to work with LibreOffice.

To check of the virtual environment is set for LibreOffice use the following command.

```python
>>> python -m main env -u
UNO Environment
```

Newer versions of [poetry] will not work with the configuration set up for LibreOffice.

When you need to use [poetry] just toggle environment.

```python
python -m main env -t
```

### Debugging Macros

Now it is possible to also debug macros when running in the Codesapce (Development container).

See [Debug Macros in Vs Code](https://github.com/Amourspirit/live-libreoffice-python/wiki/Debug-Macros-in-Vs-Code) Guide for [Live LibreOffice Python]. The guide uses Port `3002` This container uses Port `3004`.

### Testing Virtual Environment

For a quick test of environment import `uno` If there is no import error you should be good to go.

```txt
(.venv) PS C:\python_ooo_dev_tools> python
Python 3.8.10 (default, Mar 23 2022, 15:43:48) [MSC v.1928 64 bit (AMD64)] on win32
Type "help", "copyright", "credits" or "license" for more information.
>>> import uno
>>>
```

[ooouno]: https://pypi.org/project/ooouno/
[oooscript]: https://pypi.org/project/oooscript/
[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/
[OooDev]: https://python-ooo-dev-tools.readthedocs.io/en/latest/
[ooouno]: https://github.com/Amourspirit/python-ooouno
[ScriptForge]: https://gitlab.com/LibreOfficiant/scriptforge
[ex]: ./ex/
[LibreOffice Programming]: https://github.com/flywire/lo-p
[poetry]: https://python-poetry.org
[Live LibreOffice Python]:https://github.com/Amourspirit/live-libreoffice-python
