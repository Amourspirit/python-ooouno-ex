# Cell Controls

<p align="center">
<img src="https://github.com/Amourspirit/python-ooouno-ex/assets/4193389/ca31167a-319b-4320-ba93-25147e062633" width="650" height="494">
</p>

## Overview

This example takes advantage of the [OOO Development Tools] (OooDev) to demonstrate how to insert controls dynamically into cells and monitor cell change events.

Demonstrates the following:

- Automatically insert controls into cells.
- Monitor Sheet Change Events
- Monitor Cell Change Events
- Dynamically add and remove controls to Calc cells.
- Embedding a script into a Calc document.

The outputted document will dynamically insert a Currency control when a cell in the `B` column is selected and a Number control when a cell in the `C` column is selected.

The Start and Stop buttons are added to the document to start and stop the automation.


## Steps

This example takes an existing Calc Document and adds controls to cells. The controls are linked to a script that take various actions.
The script for managing the controls is embedded into the Calc document via a `Make` script.

Running `make build` in the current folder will embed the `macro_listener.py` script into the Calc `./data/src_doc/src_doc.ods` document and save it as `./data/odev_cell_controls.ods`.

Next running `python -m start` will open the document and start the automation.
The `start.py` script calls the `sheet_controls.SheetControls` class that will automatically insert button controls into the rows of the document and add the start and stop buttons to the document.
Once `start.py` has been run the document can be saved as a standalone document and the controls will still work with [OOO Development Tools Extension] installed.
This step is demonstrate how to dynamically add controls to a Calc document and register them with the `macro_listener.py` script.

Note that it is important to have the buttons use the `assign_script()` method due to a LibreOffice bug [Forms Listeners stop working after a different sheet is activated](https://bugs.documentfoundation.org/show_bug.cgi?id=159134).
If the buttons are not assigned to the script, the buttons will not work after a different sheet is activated due this bug.
A workaround this bug is to not have the controls present when the sheet is changed.

This is why the sheet activation is monitored in the `macro_listener.py` file.

```python
def on_active_sheet_changed(source: Any, event_args: EventArgs, *args, **kwargs) -> None:

    global _CURRENT_CELL, _SHEET_INDEX, _PREV_CELL
    print("Active Sheet Changed")
    doc = get_current_doc()
    try:
        if _CURRENT_CELL is not None:
            event = cast("ActivationEvent", event_args.event_data)
            sheet = doc.sheets.get_sheet(event.ActiveSheet)
            if sheet.sheet_index == _SHEET_INDEX:
                cell = sheet[_CURRENT_CELL]
                _remove_sold_ctl(sheet, cell)
                _remove_cost_ctl(sheet, cell)
    except Exception as e:
        print(f"  {e}")
```

## Documents

- The `./data/src_doc/src_doc.ods` document is the source document that is used to create the `./data/odev_cell_controls.ods` document via the `make build` command.
- The `./data/odev_cell_with_controls.ods` document an example of the final output of the automation and can be used as a standalone document with the [OOO Development Tools Extension] installed.


## Scripts

- `./macro_listener.py` - The script that is embedded into the Calc document that listens for cell changes and controls.
- `./sheet_controls.py` - The script that is used to add controls to the Calc document.
- `./start.py` - The script that is used to start the automation.

## Embedding Scripts

The `macro_listener.py` script is embedded into the Calc document via the `make build` command.

```make
oooscript compile --embed --config "$(PWD)/config.json" --embed-doc "$(PWD)/data/src_doc/src_doc.ods" --build-dir "$(PWD)/data"
```

See [OooScript Docs](https://oooscript.readthedocs.io/en/latest/) for more information on embedding scripts.

## Other Notes

The `make build` command uses [oooscript] to embed the `macro_listener.py` script into the `./data/src_doc/src_doc.ods` document and save it as `./data/odev_cell_controls.ods`.


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
python ./ex/auto/calc/odev_cell_controls/start.py
```

### Windows

```ps
python .\ex\auto\calc\odev_cell_controls\start.py
```

## Live LibreOffice Python

Instructions to run this example in [Live-LibreOffice-Python](https://github.com/Amourspirit/live-libreoffice-python).

Start Live-LibreOffice-Python in a Codespace or in a Dev Container.

In the terminal run:

```bash
cd examples
gitget 'https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/calc/odev_cell_controls'
```

This will copy the `odev_cell_controls` example to the examples folder.

In the terminal run:

```bash
cd odev_cell_controls
python -m start
```


[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/
[OOO Development Tools Extension]: https://extensions.libreoffice.org/en/extensions/show/41700
[oooscript]: https://pypi.org/project/oooscript/