# LibreOffice Python UNO Examples

Examples for [ooouno], [OOO Development Tools], [ScriptForge] projects and more.

[ooouno](https://github.com/Amourspirit/python-ooouno) project has made it possible to have a much more
user friendly way of programming for [LibreOffice](https://www.libreoffice.org/).

The goal of this project is to have ever expanding practical examples for programming
using python in [LibreOffice](https://www.libreoffice.org/) in a truly [OOP](https://en.wikipedia.org/wiki/Object-oriented_programming) way.

At this time the example can be found in the [ex].
As new example are create/submitted the example [ex] folder will be expanded.

This project uses [oooscript] to compile multiple scripts into a single script and embed output into a document.
This makes it easy to use as a LibreOffice macro.

Another fantastic resource is [LibreOffice Programming],
great documentation and many java examples.

Work in progress... more to come.

## I NEED A BREAK ...

Checkout the LibreOffice Calc [Sudoku example](./ex/calc/sudoku).

![calc_sudoku](https://user-images.githubusercontent.com/4193389/165391098-883a7647-5fc8-47de-b028-4c2c98337abe.png)

## Project Installation

This project use a virtual environment for development purposes.

[Poetry](https://python-poetry.org) is required to install this project.

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

Install requirement using Poetry.

```sh
poetry install
```

Add the `uno.py` and `unohelper.py` links to virtual environment.

```sh
python -m main cmd-link --add
```

### Windows

Windows is a little trickery. Creating a link to `uno.py` and importing it will not work as it does in Linux. This is due to the how LibreOffice implements the python environment on Windows.

The work around on Windows is a slight hack to the virtual environment.

Set up virtual environment if not existing (recommend using PowerShell).

```ps
python -m venv .\.venv
```

Get LibreOffice python version.

```ps
PS C:\python_ooo_dev_tools> & 'C:\Program Files\LibreOffice\program\python.exe'
Python 3.8.10 (default, Mar 23 2022, 15:43:48) [MSC v.1928 64 bit (AMD64)] on win32
Type "help", "copyright", "credits" or "license" for more information.
>>>
```

Edit `.\.venv/pyvenv.cfg` file.

```ps
PS C:\python_ooo_dev_tools> notepad .\.venv\pyvenv.cfg
```

Original may look something like:

```ps
home = C:\ProgramData\Miniconda3
include-system-site-packages = false
version = 3.9.7
```

Change to: With the version that is the same as current LibreOffice Version

```txt
home = C:\Program Files\LibreOffice\program
include-system-site-packages = false
version = 3.8.10
```

Activate virtual environment and install development requirements.

```ps
. .\.venv\Scripts\activate
```

```ps
poetry install
```

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
[ScriptForge]: https://gitlab.com/LibreOfficiant/scriptforge
[ex]: ./ex/
[LibreOffice Programming]: https://github.com/flywire/lo-p
