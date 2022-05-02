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

Build will compile the python scripts for this example into a single python script.

The following command will compile script as `inputbox.py` and embed it into `inputbox.odt`
The output is written into `build` folder in the projects root.

```sh
python -m main build -e --config 'ex/general/input_box/config.json' --embed-src 'ex/general/input_box/inputbox.odt'
```

## Source

see [inputbox.py](inputbox.py)
