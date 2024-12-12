from mvc.controller import MultiSyntaxController
from mvc.model import MultiSyntaxModel
from mvc.view import MultiSyntaxView

# from ooodev.loader import Lo


def show_tab_dialog(*args, **kwargs):
    # _ = Lo.current_doc
    dlg = MultiSyntaxController(model=MultiSyntaxModel(), view=MultiSyntaxView())
    dlg.start()
