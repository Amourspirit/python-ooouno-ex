# Dialog Menu Example

<p align="center">
<img src="https://github.com/Amourspirit/python-ooouno-ex/assets/4193389/63fc2b56-9f1c-47ab-9d17-d9a59d96ef25" width="600" height="606">
</p>

## Introduction

Example of added a menu to a dialog.
This demo uses [OOO Development Tools] (OooDev).

This example demonstrates how to create a Dialog window programmatically with a menu.
The dialog contains a simple ☰ icon that when moused over will display a menu.

Adding a menu bar to a dialog is not possible with the standard LibreOffice API.
Howerver, adding a menu to a dialog is possible with a bit of custom work.

The example uses `OooDev` to create a dialog with a menu. `OooDev` has great support for creating menus.

The simple explanation is that the dialog is created, a label is added with the ☰ icon, and a menu is created, and the label mouse events triggers the display of the menu. A listener for the menu is added to handle the menu item selection.

## Library files

### dialog_menu_lib module

The `dialog_menu_lib` folder contains the library files for the example.

The `menu_data` module contains the menu data and creates the popup menu via the `get_popup_menu()` method.
There is also a `get_popup_from_json()` that can be used to create a popup menu from a json file.

### menu_dialog module

The `menu_dialog` module contains the dialog class and the dialog event listener class and menu items.

The `MenuDialog` class creates the dialog and the menu items.

![Menu Dialog Screenshot](https://github.com/Amourspirit/python-ooouno-ex/assets/4193389/f5f1c785-5c07-4f6b-8290-e648ce01e5de)

When menu items are selected info is written to the dialog text area.

Some of the menu items take actions such as `File -> Close OK`, `File -> Close` and `Help -> About`.

In your own code you can choose to take any action you want when a menu item is selected.

## Other Notes

When a popup menu is created, its menu items are not executed by default when they are selected.
This actually makes for more flexibility as you can choose to execute the menu item or take some other action when the menu item is selected.

In the example, only the menu `.uno:About` and `.uno:HelpIndex` item commands are executed when selected.
The `.uno:exit` and `.uno:exitok` menu items are used to close the dialog.

```python
from ooodev.gui.menu.popup_menu import PopupMenu
if TYPE_CHECKING:
    from com.sun.star.awt import MenuEvent
# ...

class MenuDialog:

    def __init__(self) -> None:
        # ...
        self._execute_cmds = {".uno:About", ".uno:HelpIndex"}

    def on_menu_select(self, src: Any, event: EventArgs, menu: PopupMenu) -> None:
        print("Menu Selected")
        me = cast("MenuEvent", event.event_data)
        command = menu.get_command(me.MenuId)
        self._write_line(f"Menu Selected: {command}, Menu ID: {me.MenuId}")
        if command == ".uno:exit":
            self._dialog.end_execute()
            return
        if command == ".uno:exitok":
            self._dialog.end_dialog(MessageBoxResultsEnum.OK.value)
            return
        if command in self._execute_cmds and menu.is_dispatch_cmd(command):
            menu.execute_cmd(command)
```

## Macro Code

The `macro_code` module is a simple module that has the purpose of displaying  the dialog at a macro level.

```python
from dialog_menu_lib.menu_dialog import MenuDialog

def show_dialog(*args):
    dialog = MenuDialog()
    dialog.show()
```

## Make

There is a `MakeFile` in the root of this project.
By running `make` from this folder a `data/dialog_menu.ods` file is created.
This file contains the code and macro needed to run the example providing the [OooDev Extension] is installed.
The existing file is overwritten. Note the existing file already has the code and macro.

```sh
make build
```

See [Guide on embedding python macros in a LibreOffice Document](https://python-ooo-dev-tools.readthedocs.io/en/latest/guide/embed_python.html)

## start.py

The `start.py` file is a way to run the example code when this project is run in a development container.

## Example File

The Example file is located in `data/dialog_menu.ods` it is ready to run providing the [OooDev Extension] is installed.




## See Also:

- [OOO Development Tools]
- [OooDev Help Docs - Menus](https://python-ooo-dev-tools.readthedocs.io/en/latest/help/common/gui/menus/index.html)
- [OooDev Extension]
- [Guide on embedding python macros in a LibreOffice Document](https://python-ooo-dev-tools.readthedocs.io/en/latest/guide/embed_python.html)

## Automate


### Dev Container

From this folder.

```sh
python -m start
```

### Cross Platform

From this folder.

```sh
python -m start
```

### Linux/Mac

```sh
python ./ex/dialog/dialog_menu/start.py
```

### Windows

```ps
python .\ex\dialog\dialog_menu\start.py
```

## Live LibreOffice Python

Instructions to run this example in [Live-LibreOffice-Python](https://github.com/Amourspirit/live-libreoffice-python).

Start Live-LibreOffice-Python in a Codespace or in a Dev Container.

In the terminal run:

```bash
cd examples
gitget 'https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/dialog/dialog_menu'
```

This will copy the `dialog_menu` example to the examples folder.

In the terminal run:

```bash
cd dialog_menu
python -m start
```


[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/
[OooDev Extension]: https://extensions.libreoffice.org/en/extensions/show/41700