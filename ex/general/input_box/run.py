from __future__ import annotations, unicode_literals
import os
from ooodev.utils.lo import Lo
from ooodev.office.write import Write
from ooodev.utils.gui import GUI
from ooodev.conn.connectors import ConnectSocket
from script import input_box

def main(*args, **kwargs):
    os.environ["ODEV_MACRO_LOADER_OVERRIDE"] = "1"
    with Lo.Loader(connector=ConnectSocket()):
        doc = Write.create_doc()
        GUI.set_visible(visible=True, doc=doc)
        input_box()

if __name__ == "__main__":
    main()