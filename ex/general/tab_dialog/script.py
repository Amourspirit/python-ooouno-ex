from ooodev.utils.lo import Lo
from ex.general.tab_dialog.mvc.controller import MultiSyntaxController
from ex.general.tab_dialog.mvc.model import MultiSyntaxModel
from ex.general.tab_dialog.mvc.view import MultiSyntaxView
from ooodev.macro.macro_loader import MacroLoader


def show_tab_dialog(*args, **kwargs):
    with MacroLoader():
        dlg = MultiSyntaxController(model=MultiSyntaxModel(), view=MultiSyntaxView())
        dlg.start()
