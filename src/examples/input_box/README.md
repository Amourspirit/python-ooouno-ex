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

## Source

see [inputbox.py](inputbox.py)
