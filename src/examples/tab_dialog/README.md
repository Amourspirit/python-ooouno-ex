# TAB CONTROL DIALOG BOX

This sampale demonstrated a Tab Control in a Dialog Box.

Also this example is created in MVC (model, view controller) style.

Example also demonstrates usage of Radio Button Controls and List Box controls.

![Dialog](./img/tab_dialog_lite_dark.png)

## Sample Document

see sample LibreOffice Writer document, [tab_dialog.odt](tab_dialog.odt)

### Sample Code

see [script.py](script.py) for sample usage.

## Usage

```python
from src.examples.tab_dialog.mvc.controller import MultiSyntaxController
from src.examples.tab_dialog.mvc.model import MultiSyntaxModel
from src.examples.tab_dialog.mvc.view import MultiSyntaxView


dlg = MultiSyntaxController(model=MultiSyntaxModel(), view=MultiSyntaxView())
dlg.start()

```

## Source

see [mvc](mvc)
