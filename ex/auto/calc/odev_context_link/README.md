# Intercept Context Menu

<p align="center">
<img src="https://github.com/Amourspirit/python-ooouno-ex/assets/4193389/d4061112-ef0a-47b6-83b2-f02bc89051dd" width="600" height="562">
</p>

## Introduction

Example of intercepting a context menu in LibreOffice Calc.

This demo uses This demo uses [OOO Development Tools] (OooDev).

Intercepting a menu with a custom action is a powerful feature. This example demonstrates how to intercept the context menu in LibreOffice Calc and add a custom action to the menu.

It is relatively simple to add a custom action to the context menu in LibreOffice Calc.
This can be seen in the [About Example for App Menu](https://python-ooo-dev-tools.readthedocs.io/en/latest/help/common/gui/menus/app_menu/about_example.html) of the [OOO Development Tools] documentation.

Adding a custom action is a more complex task. One possible solution is to point the context menu to run a macro. A macro url could be constructed using [MacroScript.get_url_script()](https://python-ooo-dev-tools.readthedocs.io/en/latest/src/macro/script/macro_script.html#ooodev.macro.script.MacroScript.get_url_script) or [Shortcuts.get_url_script()](https://python-ooo-dev-tools.readthedocs.io/en/latest/src/gui/menu/shortcuts.html#ooodev.gui.menu.Shortcuts.get_url_script).
This solution can work in many case but has some limitations. One limitation is it can not be used to pass args to the macro.

This example demonstrates how to add a custom action to the context menu in LibreOffice Calc.
One benefit of this approach is that it can be used to pass args to the dispatch.

The example intercepts the menu for a cell in LibreOffice Calc and adds a custom `Convert to URL` action to the context menu if the cell text matches a url string but is not currently a url object.

If the cell text is already a url object, the `Convert to URL` action not inserted into the menu.


## Library files

The `odev_context_link_lib` folder contains the library files for the example.

### DispatchConvertCellUrl class

The `DispatchConvertCellUrl` class has the purpose of converting a cell text to a url object.
Its `dispatch` method get the doc from the current context. It uses the sheet and cell that was read in from the url parameters and converts the cell text to a url object.

```python
def dispatch(self, url: URL, args: Tuple[PropertyValue, ...]) -> None:
    doc = CalcDoc.from_current_doc()
    sheet = doc.sheets[self._sheet]
    cell = sheet[self._cell]
    self._convert_to_hyperlink(cell)
```

### DispatchProviderInterceptor class

The `DispatchProviderInterceptor` class has the purpose of listening to dispatch calls and intercepting the context menu dispatch call when it matches `.uno:ooodev.calc.menu.convert.url`.
When matched the args are read and then a `DispatchConvertCellUrl` instance is constructed and returned.

```python
def queryDispatch(self, url: URL, target_frame_name: str, search_flags: int):
        # ...

        if url.Main == ".uno:ooodev.calc.menu.convert.url":
            with contextlib.suppress(Exception):
                args = self._convert_query_to_dict(url.Arguments)
                return DispatchConvertCellUrl(sheet=args["sheet"], cell=args["cell"])
        # ...
```

Note that the `DispatchProviderInterceptor` class is a singleton. This is because an instance must be kept alive to work correctly.

### dispatch_mgr module


The `dispatch_mgr` module has the purpose registering and un-registering the `DispatchProviderInterceptor` instance. It also contains the `on_menu_intercept()` method that is called when the context menu is intercepted.


```python
def on_menu_intercept(
    src: ContextMenuInterceptor,
    event: EventArgsGeneric[ContextMenuInterceptorEventData],
    view: CalcSheetView,
) -> None:
    with contextlib.suppress(Exception):

        container = event.event_data.event.action_trigger_container
        event.event_data.action = ContextMenuAction.CONTINUE_MODIFIED
        if (
            container[0].CommandURL == ".uno:Cut"
            and container[-1].CommandURL == ".uno:FormatCellDialog"
        ):
            selection = event.event_data.event.selection.get_selection()

            if selection.getImplementationName() == "ScCellObj":
                addr = cast("CellAddress", selection.getCellAddress())
                doc = CalcDoc.from_current_doc()
                sheet = doc.get_active_sheet()
                cell_obj = doc.range_converter.get_cell_obj_from_addr(addr)
                cell = sheet[cell_obj]
                if not is_http_url(cell.get_string()):
                    return
                if not has_url(cell):
                    container.insert_by_index(
                        4,
                        ActionTriggerItem(
                            f".uno:ooodev.calc.menu.convert.url?sheet={sheet.name}&cell={cell_obj}",
                            "Convert to URL"
                        )
                    )
                    container.insert_by_index(4, ActionTriggerSep())
                    event.event_data.action = ContextMenuAction.CONTINUE_MODIFIED
```


## Macro Code

The `macro_code` module is a simple module that has the purpose of registering and unregistering at a macro level.
These method are used to register and unregister the methods and events when the document is loaded and unloaded.

```python
from ooodev.calc import CalcDoc
from odev_context_link_lib import dispatch_mgr

def register_url_interceptor(*args):
    doc = CalcDoc.from_current_doc()
    dispatch_mgr.register_interceptor(doc)

def unregister_url_interceptor(*args):
    doc = CalcDoc.from_current_doc()
    dispatch_mgr.unregister_interceptor(doc)
```

## Make

There is a `MakeFile` in the root of this project.
By running `make` from this folder a `data/links.ods` file is created.
This file contains the code and macro needed to run the example providing the [OooDev Extension] is installed.
The existing file is overwritten. Note the existing file already has the code and macro.

```sh
make build
```

See [Guide on embedding python macros in a LibreOffice Document](https://python-ooo-dev-tools.readthedocs.io/en/latest/guide/embed_python.html)

## start.py

The `start.py` file is a way to run the example code when this project is run in a development container.

## Example File

The Example file is located in `data/links.ods` it is ready to run providing the [OooDev Extension] is installed.
The file has a start up macro that registers the `DispatchProviderInterceptor` instance.
Simply open the file and the intercept will be active.

### Note

If you are having trouble getting the example to work, try turning off `Automatic Spell checking` in the `Tools` menu (`Shift+F7`).

## Automate

A message box is display once the document has been processed asking if you want to close the document.

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
python ./ex/auto/calc/odev_context_link/start.py
```

### Windows

```ps
python .\ex\auto\calc\odev_context_link\start.py
```

## Live LibreOffice Python

Instructions to run this example in [Live-LibreOffice-Python](https://github.com/Amourspirit/live-libreoffice-python).

Start Live-LibreOffice-Python in a Codespace or in a Dev Container.

In the terminal run:

```bash
cd examples
gitget 'https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/calc/odev_context_link'
```

This will copy the `odev_context_link` example to the examples folder.

In the terminal run:

```bash
cd odev_context_link
python -m start
```


[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/
[OooDev Extension]: https://extensions.libreoffice.org/en/extensions/show/41700
