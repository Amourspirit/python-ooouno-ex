from src.examples.tab_dialog.mvc.controller import MultiSyntaxController
from src.examples.tab_dialog.mvc.model import MultiSyntaxModel
from src.examples.tab_dialog.mvc.view import MultiSyntaxView


def show_tab_dialog(*args, **kwargs):
    dlg = MultiSyntaxController(model=MultiSyntaxModel(), view=MultiSyntaxView())
    dlg.start()
