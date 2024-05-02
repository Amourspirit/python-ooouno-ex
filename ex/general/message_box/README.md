# MESSAGE BOX

Message box is a dialog box give user feedback.

See Also: [OooDev Message Box](https://python-ooo-dev-tools.readthedocs.io/en/latest/src/dialog/msgbox.html)

## Sample document

See sample LibreOffice Writer document, [msgbox.odt](msgbox.odt)

### sample code

see [script.py](script.py) for sample usage.

## usage

### Simple user info

```python
from ooodev.dialog.msgbox import (
    MsgBox,  MessageBoxButtonsEnum, MessageBoxType
)

results = MsgBox.msgbox(
    msg="Hello World",
    title="Nice Day!",
    buttons=MessageBoxButtonsEnum.BUTTONS_OK,
    boxtype=MessageBoxType.INFOBOX,
)
```

### Ask user Yes or No

```python
from ooodev.dialog.msgbox import (
    MsgBox,  MessageBoxButtonsEnum, MessageBoxResultsEnum, MessageBoxType
)

result = MsgBox.msgbox(
        msg="Are you sure you want to by coffee?",
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

See [Guide on embedding python macros in a LibreOffice Document](https://python-ooo-dev-tools.readthedocs.io/en/latest/guide/embed_python.html).

## Run Directly

To start LibreOffice and display a message box run the following command from this folder.

```sh
make msg-short
```

or

```sh
make msg-long
```

or

```sh
make msg-warn
```

or

```sh
make msg-error
```

or

```sh
make msg-yes
```

