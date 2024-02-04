# This example depends on code outside the this source folder.
# For this reason the project path is inserted into sys.path
#
# The tab dialog uses the ooodev library, therefore, MacroLoader is used to load the library context in the macro methods.
# In this script we don't want the MacroLoader to reset the context, so we set the environment variable ODEV_MACRO_LOADER_OVERRIDE to 1.
#
# Lo.Loader() context manager is used to start LibreOffice and close it when the context manager exits.
from __future__ import annotations, unicode_literals
import os
import sys


def _insert_project_path() -> None:
    """
    Inserts project path into sys.path
    """
    path = os.path.dirname(os.path.abspath(__file__))

    while not os.path.exists(os.path.join(path, ".root_token")):
        path = os.path.dirname(path)
        if path == "/":
            raise FileNotFoundError(
                ".root_token file not found in any parent directory"
            )
    sys.path.append(path)


_insert_project_path()


from ooodev.loader import Lo
from ooodev.office.write import Write
from ooodev.utils.gui import GUI
from ooodev.conn.connectors import ConnectSocket
from script import show_tab_dialog


def main(*args, **kwargs):
    os.environ["ODEV_MACRO_LOADER_OVERRIDE"] = "1"
    with Lo.Loader(connector=ConnectSocket()):
        doc = Write.create_doc()  # create a new document for Context
        GUI.set_visible(visible=True, doc=doc)  # make the document visible
        show_tab_dialog()  # show the tab dialog over the document
        assert True


if __name__ == "__main__":
    main()
