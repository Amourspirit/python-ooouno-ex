#!/usr/bin/env python
"""
This module is not to be called directly
and is intended to be called from the projects main.
Such as: python -m main auto --process "ex/auto/writer/hello_world/start.py"
"""
from __future__ import annotations
import scriptforge as SF
from typing import TYPE_CHECKING
from ooo.dyn.beans.property_value import PropertyValue
from src.lib import writer_sel_framework as o_sel
from src.lib.connect import LoSocketStart

if TYPE_CHECKING:
    from com.sun.star.text import XText
    from com.sun.star.frame import DispatchHelper


def main() -> int:
    # set the port to start soffice on
    port = 2002
    # create an instance used to launch soffice as a server
    lo = LoSocketStart(port=port, start_soffice=True)

    # launch soffice (LibreOffice server)
    lo.connect()

    # Connect ScritForge to the running soffice server
    SF.ScriptForge(hostname=lo.host, port=lo.port)

    # connect to ui
    ui: SF.SFScriptForge.SF_UI = SF.CreateScriptService("UI")
    bas: SF.SFScriptForge.SF_Basic = SF.CreateScriptService("Basic")

    # create a new Writer document
    doc: SF.SFDocuments.SF_Writer = ui.CreateDocument("Writer")
    odoc: XText = doc.XComponent

    # get XText t can access to text cursor
    txt = odoc.getText()

    # get document cursor
    cursor = txt.getEnd()

    # write text
    cursor.setString("Hello World")

    bas: SF.SFScriptForge.SF_Basic = SF.CreateScriptService("Basic")

    # create a dispatch helper
    dispatcher: DispatchHelper = bas.CreateUnoService(
        "com.sun.star.frame.DispatchHelper"
    )

    # create style properties
    style = "Title"  # represents Title style of document
    args = (
        PropertyValue(Name="Template", Value=style),
        PropertyValue(Name="Family", Value=2),
    )
    # apply style to selected text
    dispatcher.executeDispatch(
        bas.ThisComponent.CurrentController.Frame, ".uno:StyleApply", "", 0, args
    )

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
