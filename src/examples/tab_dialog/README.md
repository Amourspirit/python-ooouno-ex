# TAB CONTROL DIALOG BOX

This sampale demonstrated a Tab Control in a Dialog Box.

Also this example is created in MVC (model, view controller) style.

## sample

## sample document

see sample LibreOffice Writer document, [dialog_tab.odt](dialog_tab.odt)

### sample code

see [script.py](script.py) for sample usage.

## usage

```python
from src.examples.tab_dialog.mvc.controller import MultiSyntaxController
from src.examples.tab_dialog.mvc.model import MultiSyntaxModel
from src.examples.tab_dialog.mvc.view import MultiSyntaxView


dlg = MultiSyntaxController(model=MultiSyntaxModel(), view=MultiSyntaxView())
dlg.start()

```

## source

see [mvc](mvc)
