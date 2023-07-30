from __future__ import annotations, unicode_literals
import os
import sys
from ooodev.utils.lo import Lo
from ooodev.office.write import Write
from ooodev.utils.gui import GUI
from ooodev.conn.connectors import ConnectSocket
import script

def main(arg: str = ""):
    os.environ["ODEV_MACRO_LOADER_OVERRIDE"] = "1"
    with Lo.Loader(connector=ConnectSocket()):
        doc = Write.create_doc()
        GUI.set_visible(visible=True, doc=doc)
        
        if arg == "short":
            script.msg_small()
        elif arg == "long":
            script.msg_long()
        elif arg == "warn":
            script.msg_warning()
        elif arg == "error":
            script.msg_error()
        else:
            script.msg_default_yes()

if __name__ == "__main__":
    if len(sys.argv) > 1:
        main(sys.argv[1])
    else:
        main()