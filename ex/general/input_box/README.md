# INPUT BOX

Input box is a dialog box that ask for user input.

## Sample Document

See sample LibreOffice Writer document, [inputbox.odt](inputbox.odt)

### Sample Code

see [script.py](script.py) for sample usage.

## usage

```python
from src.examples.input_box.inputbox import InputBox

box = InputBox("")
    result: str = box.show()
    if len(result) == 0:
        ...
        # process no input
    else:
        ...
        # process input
```

## Build

For automatic build run the following command from this folder.

```sh
make build
```

The following instructions are for manual build.


Build will compile the python scripts for this example into a single python script.

The following command will compile script as `inputbox.py` and embed it into `inputbox.odt`
The output is written into `build/input_box` folder in the projects root.

```sh
oooscript compile --embed --config "ex/general/input_box/config.json" --embed-doc "ex/general/input_box/inputbox.odt" --build-dir "build/input_box"
```

## Source

see [inputbox.py](inputbox.py)
