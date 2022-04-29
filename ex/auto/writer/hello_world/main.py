#!/usr/bin/env python
import scriptforge as SF
import sys
import os
from pathlib import Path
def _set_root_env():
    root = os.environ.get("project_root", None)
    if root is None:
        sys.path.insert(0, str(Path(__file__).parent.parent.parent.parent.parent))
_set_root_env()
from src.lib.connect import LoManager

def main():
    port = 2002
    with LoManager(use_pipe=False, port=port) as _:
        SF.ScriptForge(hostname='localhost', port=port)
        ui: SF.SFScriptForge.SF_UI = SF.CreateScriptService("UI")
        bas: SF.SFScriptForge.SF_Basic = SF.CreateScriptService("Basic")
        ui.CreateDocument("Writer")
        # doc = ui.OpenDocument("~/Documents/myFile.ods")
        bas.MsgBox("Hello World")

if __name__ == "__main__":
    main()