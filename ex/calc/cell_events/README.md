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

[ScriptForge]: https://gitlab.com/LibreOfficiant/scriptforge
