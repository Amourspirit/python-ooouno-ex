# Dialog Counter

Simple [ScriptForge] example of Dialog counter using a dialog that is embedded in Calc document.

## Sample Document

see sample LibreOffice Writer document, [BasicCounter.ods](BasicCounter.ods)

### Source

see [counter_box.py](./counter_box.py)

## Build

For automatic build run the following command from this folder.

```sh
make build
```
The following instructions are for manual build.

Build will compile the python scripts for this example into a single python script.

The following command will compile script as `BasicCounter.py` and embed it into`BasicCounter.ods`
The output is written into `build/BasicCounter` folder in the projects root.

```sh
oooscript compile --embed --config "ex/general/dialog_basic_counter/config.json" --embed-doc "ex/general/dialog_basic_counter/BasicCounter.ods" --build-dir "build/BasicCounter"
```

![screenshot](https://user-images.githubusercontent.com/4193389/179670709-978fd704-db5e-4225-ae65-92bba0e88ac8.png)


[ScriptForge]: https://gitlab.com/LibreOfficiant/scriptforge
