# Load Form from Database and raise events


Demonstrates opening a from from a database and raising events.

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
Notify Event:
OnLayoutFinished
Form with name "MainForm" Loaded.
Closing the document
Notify Event:
OnViewClosed 
Notify Event:
OnUnload     
Notify Event:
OnUnfocus
Closing Office
Office terminated
Office bridge has gone!!
```