from ex.general.tab_dialog.mvc.controller import MultiSyntaxController
from ex.general.tab_dialog.mvc.model import MultiSyntaxModel
from ex.general.tab_dialog.mvc.view import MultiSyntaxView


def show_tab_dialog(*args, **kwargs):
    dlg = MultiSyntaxController(model=MultiSyntaxModel(), view=MultiSyntaxView())
    dlg.start()
