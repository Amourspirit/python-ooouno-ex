# Load Form from Database and raise events

Demonstrates opening a from from a database and raising events.

The `XEventListener` interface is used to listen for event to be notified when a database form is loaded.

In this example the `XEventListener` interface is implemented by the `DocumentEventListener` class.

The other Event Listening Classes are not needed for this example but are included for educational purposes.

This a simple example only.

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

## Output

```text
Loading Office...
Opening D:\Users\user\Projects\python-ooouno-ex\resources\odb\Example_Sport.odb
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