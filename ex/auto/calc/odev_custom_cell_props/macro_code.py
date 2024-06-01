from ooodev.calc import CalcDoc
from ooodev.calc.cell.custom_prop_clean import CustomPropClean
from odev_custom_cell_props_lib import dispatch_mgr


def register_custom_prop_interceptor(*args):
    doc = CalcDoc.from_current_doc()
    dispatch_mgr.register_interceptor(doc)


def unregister_custom_prop_interceptor(*args):
    doc = CalcDoc.from_current_doc()
    dispatch_mgr.unregister_interceptor(doc)


def clean_custom_props(*args):
    """
    Cleans up any custom property artifacts.
    This is not critical but can be useful to keep the document clean.
    When cell are copied or deleted it may leave behind custom property artifacts.
    If these artifacts are not cleaned up, it will not interfere with the document.
    This a rather minor issue. But this function can be used to clean up the document.
    """
    doc = CalcDoc.from_current_doc()
    for sheet in doc.sheets:
        CustomPropClean(sheet).clean()
