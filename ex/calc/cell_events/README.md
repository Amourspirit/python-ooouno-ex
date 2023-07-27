# Calc Cell Events Example

This example demonstrates **Content Changed** events and **Selection Changed** events using [ScriptForge].

## Sample Document

See sample LibreOffice Writer document [cell_events.ods](cell_events.ods).

## Build

Build will compile the python scripts for this example into a single python script.

The following command will compile script as `script.py` and embed it into `cell_events.ods`
The output is written into `build` folder in the projects root.

```sh
oooscript compile --embed --config "ex/calc/cell_events/config.json" --embed-doc "ex/calc/cell_events/cell_events.ods"
```

![calc_on_sel_change](https://user-images.githubusercontent.com/4193389/166338567-e597c1e9-854c-4254-bbf8-fb8f94598797.gif)

## Live LibreOffice Python

Instructions to run this example in [Live-LibreOffice-Python](https://github.com/Amourspirit/live-libreoffice-python).

Start Live-LibreOffice-Python in a Codespace or in a Dev Container.

In the terminal run:

```bash
cd macro
gitget 'https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/calc/cell_events'
```

This will copy the `cell_events` example to the macro folder.

[ScriptForge]: https://gitlab.com/LibreOfficiant/scriptforge
