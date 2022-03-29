# MESSAGE BOX

Message box is a dialog box give user feedback.

## sample

## sample document

see sample LibreOffice Writer document, [msgbox.odt](msgbox.odt)

### sample code

see [script.py](script.py) for sample usage.

## usage

### Simple user info

```python
from src.examples.message_box.msgbox import (
    msgbox,
    MessageBoxButtonsEnum,
    MessageBoxType
)

msgbox(
    message="Hello World",
    title="Nice Day!",
    buttons=MessageBoxButtonsEnum.BUTTONS_OK,
    boxtype=MessageBoxType.INFOBOX,
)
```

### Ask user Yes or No

```python
from src.examples.message_box.msgbox import (
    msgbox,
    MessageBoxButtonsEnum,
    MessageBoxResultsEnum,
    MessageBoxType
)

result = msgbox(
        message="Are you sure you want to by coffee?",
        buttons=MessageBoxButtonsEnum.BUTTONS_YES_NO,
        title="Long...",
        boxtype=MessageBoxType.QUERYBOX,
    )

if result == MessageBoxResultsEnum.YES:
    ...
    # looks like your buying coffee.

elif result == MessageBoxResultsEnum.NO:
    ...
    # I guess another day...
```

## source

see [msgbox.py](msgbox.py)
