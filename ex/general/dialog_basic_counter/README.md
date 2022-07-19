# Dialog Counter

Simple [ScriptForge] example of Dialog counter using a dialog that is embedded in Calc document.

## Sample Socument

see sample LibreOffice Writer document, [BasicCounter.ods](BasicCounter.ods)

### sample Code

see [script.py](script.py) for sample usage.



## Build

Build will compile the python scripts for this example into a single python script.

The following command will compile script as `BasicCounter.py` and embed it into`BasicCounter.ods`
The output is written into `build` folder in the projects root.

```sh
python -m main build -e --config "ex/general/dialog_basic_counter/config.json" --embed-src "ex/general/dialog_basic_counter/BasicCounter.ods"
```

## Source

see [counter_box.py](./counter_box.py)

![screenshot](https://user-images.githubusercontent.com/4193389/179668010-30a7c762-ef83-4431-a7dd-b48ebd169de2.png)

[ScriptForge]: https://gitlab.com/LibreOfficiant/scriptforge
