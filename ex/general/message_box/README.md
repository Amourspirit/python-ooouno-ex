# MESSAGE BOX

Message box is a dialog box give user feedback.

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
        title="☕ - coffee - ☕",
        boxtype=MessageBoxType.QUERYBOX,
    )

if result == MessageBoxResultsEnum.YES:
    ...
    # looks like your buying coffee.

elif result == MessageBoxResultsEnum.NO:
    ...
    # I guess another day...
```

## Build

For automatic build run the following command from this folder.

```sh
make build
```

The following instructions are for manual build.

Build will compile the python scripts for this example into a single python script.

The following command will compile script as `msgbox.py` and embed it into`msgbox.odt`
The output is written into `build/message_box` folder in the projects root.

```sh
oooscript compile --embed --config "ex/general/message_box/config.json" --embed-doc "ex/general/message_box/msgbox.odt" --build-dir "build/message_box"
```

## Source

see [msgbox.py](msgbox.py)
