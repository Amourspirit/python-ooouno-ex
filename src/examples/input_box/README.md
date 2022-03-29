# INPUT BOX

Input box is a dialog box that ask for user input.

## sample

## sample document

see [inputbox.odt](inputbox.odt)

### sample code

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

## source

see [inputbox.py](inputbox.py)
