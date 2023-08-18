<p align="center">
<img src="https://user-images.githubusercontent.com/4193389/180623026-9e5b96fc-22c0-43b8-a612-139eb3b28737.png" alt="dispatch"/>
</p>

# Dispatch Commands Example

This example demonstrates how to dispatch commands using [OOO Development Tools] (OooDev).

Also demonstrates how to create message dialog.

LibreOffice has a comprehensive web page listing all the dispatch commands [Development/DispatchCommands](https://wiki.documentfoundation.org/Development/DispatchCommands).

This example also demonstrates hooking OooDev events that in this case allow for finer control over which commands are dispatched.

See Also: [OOO Development Tools - Chapter 4. Listening, and Other Techniques](https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part1/chapter04.html)

See [source code](./start.py)

## Automate

A message box is display once the document has been processed asking if you want to close the document.

Running the following command opens a Write document, puts it into read-only mode.
Next a Get Involved wep page is opened in local web browser.

### Dev Container

From current example folder.

```sh
python -m start -d "./data/story.odt"
```

### Cross Platform

From current example folder.

```sh
python -m start -d "./data/story.odt"
```

### Linux

From project root folder

Linux/Mac

```sh
python "./ex/auto/general/odev_dispatch/start.py" -d "ex/auto/general/odev_dispatch/data/story.odt"
```

### Windows

From project root folder

```ps
python ".\ex\auto\general\odev_dispatch\start.py" -d "ex/auto/general/odev_dispatch/data/story.odt"
```

## Live LibreOffice Python

Instructions to run this example in [Live-LibreOffice-Python](https://github.com/Amourspirit/live-libreoffice-python).

Start Live-LibreOffice-Python in a Codespace or in a Dev Container.

In the terminal run:

```bash
cd examples
gitget 'https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/general/odev_dispatch'
```

This will copy the `odev_dispatch` example to the examples folder.

In the terminal run:

```bash
cd odev_dispatch
python -m start -h
```

### Example console output

```text
PS D:\Users\user\Python\python-ooouno-ex> python .\ex\auto\general\odev_dispatch\start.py -d "resources\odt\story.odt"
Loading Office...
Opening D:\Users\user\Python\python-ooouno-ex\resources\odt\story.odt
Dispatching: ReadOnlyDoc
Dispatched: ReadOnlyDoc
Dispatching: GetInvolved
Dispatched: GetInvolved
About dispatch canceled
```

[OOO Development Tools]:https://python-ooo-dev-tools.readthedocs.io/en/latest/
