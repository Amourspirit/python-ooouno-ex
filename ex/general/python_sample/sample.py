from __future__ import unicode_literals


def doc_dialog():
    """Display a doc-based dialog"""
    model = XSCRIPTCONTEXT.getDocument()
    smgr = XSCRIPTCONTEXT.getComponentContext().ServiceManager
    dp = smgr.createInstanceWithArguments("com.sun.star.awt.DialogProvider", (model,))
    dlg = dp.createDialog("vnd.sun.star.script:Standard.Dialog1?location=document")
    dlg.execute()
    dlg.dispose()


g_exportedScripts = (doc_dialog,)
