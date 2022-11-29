# Calc Modify Listener

<p align="center">
<img src="https://user-images.githubusercontent.com/4193389/204155527-4e975c63-ea78-4591-a659-d9ddafa8970c.png" width="327" height="218">
</p>

This is a basic example that shows how attach a [XModifyListener] listener to a Calc spreadsheet so that its `modified()` method will be triggered whenever a cell is changed.

As cells in spreadsheet are modified the cell value is outputted to console.

Also demos how to attach a window listener to office.

This script will stay running until office is closed or `ctl+c` is pressed unless `-t` is passed as a parameter.

This demo uses [OOO Development Tools] (ODEV).

A `main_loop()` method is called that watches until Office is closed.

See Also:

- [OOO Development Tools - Chapter 4. Listening, and Other Techniques](https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part1/chapter04.html)
- [OOO Development Tools - Chapter 25. Monitoring Sheets](https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part4/chapter25.html)
- [Calc Select Listener Example](../odev_select_listener/)
- [Office Window Listener Example](../../general/odev_listen/)

## Listeners

This example has two classes (`ModifyListener`, `ModifyListenerAdapter`) that show different ways to subscribe to Calc spreadsheet modify events.

Functionally these two classes are identical.

Both classes are responsible for starting office.

### Using ODEV TopWindowListener Class

[ODEV] has a `ModifyListener` class that raise events that can be subscribed to.

[modify_listener_adapter.py](./modify_listener_adapter.py) (`ModifyListenerAdapter` class) demonstrates how `ModifyListener` is implemented. This means there is no need to create a class that inherits from [XModifyListener]

### Implementing XModifyListener

[modify_listener.py](./modify_listener.py) (`DocWindow` class) demonstrates implementing [XModifyListener]

## Automate

### Command Line Parameters

- `-a` runs demo using `ModifyListenerAdapter` else `ModifyListener` is used.
- `-t` runs demo and automatically closes office.
- `-h` Displays help on options.

### Cross Platform

Run from this folder.

```sh
python start.py
```

### Linus/Mac

From project root folder.

```sh
python ./ex/auto/calc/odev_modify_listener/start.py
```

### Windows

From project root folder.

```ps
python .\ex\auto\calc\odev_modify_listener\start.py
```

### Example console output

User interactions with window are reflected in console window.

Starts Write as a new document and monitors window activity and auto terminates.

```ps
PS D:\Users\user\Python\python-ooouno-ex> python .\ex\auto\calc\odev_modify_listener\start.py
Press 'ctl+c' to exit script early.
Loading Office...
Creating Office document scalc
Modified
  A4 = 66.5
Modified
  A4 = 66.5
Modified
  A6 = 33.45
Modified    
  A6 = 33.45
Modified
  C6 = 35.998
Modified
  C6 = 35.998
Closing
Closing the document
Disposing
Closing Office
Office terminated

Exiting by document close.
```

[XModifyListener]: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1util_1_1XModifyListener.html

[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/
[ODEV]: https://python-ooo-dev-tools.readthedocs.io/en/latest/
