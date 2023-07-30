# Load Form from Database and raise events

Demonstrates opening a from from a database and raising events.

The `XEventListener` interface is used to listen for event to be notified when a database form is loaded.

In this example the `XEventListener` interface is implemented by the `DocumentEventListener` class.

The other Event Listening Classes are not needed for this example but are included for educational purposes.

This a simple example only.

## Dev Container

From current example folder.

```shell
python -m start
```

## Cross Platform

From current example folder.

```shell
python -m start
```

## Linux/Mac

From project root folder

```sh
python ./ex/auto/forms/form_load_event/start.py
```

## Windows

From project root folder

```ps
python .\ex\auto\forms\form_load_event\start.py
```

## Live LibreOffice Python

Instructions to run this example in [Live-LibreOffice-Python](https://github.com/Amourspirit/live-libreoffice-python).

Start Live-LibreOffice-Python in a Codespace or in a Dev Container.

In the terminal run:

```bash
cd examples
gitget 'https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/forms/form_load_event'
```

This will copy the `form_load_event` example to the examples folder.

In the terminal run:

```bash
cd form_load_event
python -m start
```

## Output

```text
Loading Office...
Opening /home/user/Projects/ooouno_ex/ex/auto/forms/form_load_event/data/Example_Sport.odb
Notify Event: OnVisAreaChanged
Notify Event: OnPageCountChange
Notify Event: OnLoad
Form with name "MainForm" Loaded.
Closing the document
Notify Event: OnLayoutFinished
Notify Event: OnViewClosed
Notify Event: OnUnload
Notify Event: OnUnfocus
Closing Office
Office terminated
Office bridge has gone!!
```