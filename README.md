# Live LibreOffice Python UNO Examples

An ever-expanding collection of practical Python examples for manipulating [LibreOffice](https://www.libreoffice.org/) components (Calc spreadsheets, Writer text documents, Impress presentation slide decks, Draw vector graphocs, etc.). A big advatage of automating LibreOffice through Python, rather than Basic, is that Python is much more expressive and grounded in first principles -- the "[OOP](https://en.wikipedia.org/wiki/Object-oriented_programming) way."

So far, this collection includes the use of the [ooouno], [OOO Development Tools] (OooDev), and [ScriptForge] projects.

[OooDev] is a powerful python library for LibreOffice and is also available as an [extension](https://extensions.libreoffice.org/en/extensions/show/41700) for LibreOffice.

[ooouno] greatly eases Python programming for [LibreOffice](https://www.libreoffice.org/). When you use [OooDev], you get [ooouno] automatically.

[ScriptForge] is a repository of "macro scripting resources" that are written in basic, but callable from python. They have been contributed for consideration to be incorporated in future distributions of LibreOffice (hence the "forge" aspect of the name.) In the mean time, we can access them explicitly via the scriptforge PyPI package. See https://gitlab.com/LibreOfficiant/scriptforge.

### TL;DR

- The main collection of examples reside in [ex], with a subfolder for each component, and with a sub-subfolder for each example. Every folder and subfolder has its own README.md file.
- Some secondary examples reside in `macro`. There are no README files for those, yet.
- Code that supports setting up the environment needed to run these examples reside in `src`, starting with `src/main.py`. There are a few corresponding unit tests in `tests`.
- The `resources` folder contains assets that are used by the examples (pre-existing ods/odb/odt files, plain text files, images, etc.) NOTE: Some of the examples have their own resources within their subfolders.

### oooscript

This project uses [oooscript] to compile multiple scripts into a single script and embed output into a document.
This makes it easy to use as a LibreOffice macro.
See the [embedding python macros in a LibreOffice Document](https://python-ooo-dev-tools.readthedocs.io/en/latest/guide/embed_python.html) guide.

### Other Resources

Another fantastic resource is [LibreOffice Programming].
It has great documentation and many Java examples.
An archive of the Java code is available at https://github.com/Amourspirit/libreoffice_lop_java.

## I NEED A BREAK ...

Check out the LibreOffice Calc [Sudoku example](./ex/calc/sudoku).

![calc_sudoku](https://user-images.githubusercontent.com/4193389/165391098-883a7647-5fc8-47de-b028-4c2c98337abe.png)

## Running the Project Examples

This project is set up to work with [Vs Code](https://code.visualstudio.com/), [Pycharm](https://www.jetbrains.com/pycharm), or [Github Codespaces](https://docs.github.com/en/codespaces/overview), or in a Development Container on your local computer.
You can run the examples directly from within any of them.
The Development container for this project is based off [Live LibreOffice Python](https://github.com/Amourspirit/live-libreoffice-python) template. See the [Live LibreOffice Python Wiki](https://github.com/Amourspirit/live-libreoffice-python/wiki) for more information.

## Project Installation

Installing this project is not necessary if you are using [Github Codespaces](https://docs.github.com/en/codespaces/overview) or in a Development Container on your local computer; however, if you want to install locally, follow the instructions below.

[poetry] is used to install this project. Poetry is one of several back-end tools available for building Python packages. It offers supperb dependency management and virtual environment management.

Although [poetry] is recommended to set up this project it is not strickly speaking required.
One could manually set the virtual envornment and install the required pacakges in the `pyproject.toml` file.

One tricky part of running these examples is gaining access to the `uno` and `unohelper` packages that are embedded in LibreOffice (alongside the embedded build of Python itself). [ooouno] depends on them and they are the glue between python scritps and LibreOffice. The exact location of these files vary depending on OS and other factors. This project uses [oooenv] which is a command line tool to aid in connecting a virtual environment to the `uno` an `unohelper` modules.

### Linux

Clone this project into an appropriate working folder, then navigate to that folder and set up virtual environment.
Activate the virtual environment and have [poetry] install the dependencies (both the runtime and development requirements).

```sh
python3 -m venv .venv
source .venv/bin/activate
poetry install
```

Add the `uno.py` and `unohelper.py` links to virtual environment using the [oooenv] tool which is installed in the virtual environment.

```sh
oooenv cmd-link -a
```

### Windows

Windows is a little tricky. Creating a link to `uno.py` and importing it will not work as it does in Linux. This is due to the how LibreOffice implements the python environment on Windows.
The workaround on Windows is a slight hack to the virtual environment.

The initial steps are basically the same as for Linux:
Clone this project into an appropriate working folder, then navigate to that folder and set up virtual environment (recommend using PowerShell).
TIP: In Windows Explorer, Shift+RightClick offers "Open PowerShell window here".
Activate the virtual environment and have [poetry] install the dependencies (both the runtime and development requirements).

```ps
py -m venv .\.venv
.\.venv\Scripts\activate
poetry install
```

Now, here's how we hack the virtual environment using the [oooenv] tool which is installed in the virtual environment.
We're creating a secondary virtual environment configuration that runs off of the copy of Python 3.8 that is embedded in LibreOfiice (where the `uno` and `unohelper` packages are installed.)
The first time you run this `oooenv env -t` command, the secondary environment configuration is created and activated.
Thereafter, the `-t` switch tells the command to toggle between the main virtual environment and this special UNO evnvironment.

```ps
oooenv env -t
```
To check if the virtual environment is set for LibreOffice use the `-u` switch.

```ps
oooenv env -u
UNO Environment
```

"Why ever toggle back?" you ask.
The UNO environment is stuck with whatever version of Python is embedded in LibreOfiice (3.8.11 at the time of this writing).
Newer versions of [poetry], for example, require 3.9 or higher.
So, when you need to use [poetry] just toggle the environment, then toggle back.

```
oooenv env -t
poetry <some-command>
oooenv env -t
```

If you get an error such as `ImportError: Are you sure that uno has been imported?` then it is likely that the `uno` environment has not been activated.
Try running `oooenv env -t`.

### Testing the Installed Environment

For a quick test to see if you correctly have the virtual environment set up to depend on the `uno` package, simply launch python and try to import `uno`.
If there is no import error, then you should be good to go.

```txt
(.venv) PS C:\python_ooo_dev_tools> python
Python 3.8.10 (default, Mar 23 2022, 15:43:48) [MSC v.1928 64 bit (AMD64)] on win32
Type "help", "copyright", "credits" or "license" for more information.
>>> import uno
>>>
```

### Debugging Macros

It is possible to debug macros when running in the Codesapce (Development container).

See [Debug Macros in Vs Code](https://github.com/Amourspirit/live-libreoffice-python/wiki/Debug-Macros-in-Vs-Code) Guide for [Live LibreOffice Python]. The guide uses Port `3002` This container uses Port `3004`.



[ooouno]: https://pypi.org/project/ooouno/
[oooscript]: https://pypi.org/project/oooscript/
[oooenv]: https://pypi.org/project/oooenv/
[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/
[OooDev]: https://python-ooo-dev-tools.readthedocs.io/en/latest/
[ooouno]: https://github.com/Amourspirit/python-ooouno
[ScriptForge]: https://gitlab.com/LibreOfficiant/scriptforge
[ex]: ./ex/
[LibreOffice Programming]: https://github.com/flywire/lo-p
[poetry]: https://python-poetry.org
[Live LibreOffice Python]:https://github.com/Amourspirit/live-libreoffice-python
