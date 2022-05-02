#!/usr/bin/env python
from __future__ import annotations
import scriptforge as SF
from typing import TYPE_CHECKING
from ooo.dyn.beans.property_value import PropertyValue
from src.lib import writer_sel_framework as o_sel
from src.lib.connect import LoSocketStart

if TYPE_CHECKING:
    from com.sun.star.text import XText
    from com.sun.star.text import XWordCursor
    from com.sun.star.frame import DispatchHelper
    from ooo.lo.text.x_text_cursor import XTextCursor

def main():
    port = 2002
    lo = LoSocketStart(port=port, start_soffice=True)
    lo.connect()
    SF.ScriptForge(hostname=lo.host, port=lo.port)
    ui: SF.SFScriptForge.SF_UI = SF.CreateScriptService("UI")
    bas: SF.SFScriptForge.SF_Basic = SF.CreateScriptService("Basic")
    doc: SF.SFDocuments.SF_Writer = ui.CreateDocument("Writer")
    odoc: XText = doc.XComponent
    txt = odoc.getText()
    cursor: XWordCursor = txt.getEnd()
    cursor.setString("Hello World")
    o_sel.select_view_by_cursors(sel=cursor, require_selection=False)
    bas: SF.SFScriptForge.SF_Basic = SF.CreateScriptService("Basic")
    dispatcher: DispatchHelper = bas.CreateUnoService('com.sun.star.frame.DispatchHelper')
    style = "Title"
    args = (
        PropertyValue(Name='Template', Value=style),
        PropertyValue(Name='Family',Value=2 )
    )
    dispatcher.executeDispatch(bas.ThisComponent.CurrentController.Frame, ".uno:StyleApply", "", 0, args)
    v_cursor: XTextCursor = bas.ThisComponent.CurrentController.Frame.getController().getViewCursor()
    v_cursor.gotoEnd(False)
    raise SystemExit

if __name__ == "__main__":
    main()