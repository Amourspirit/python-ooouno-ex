from ooodev.calc import CalcDoc
from odev_context_link_lib import dispatch_mgr


def register_url_interceptor(*args):
    doc = CalcDoc.from_current_doc()
    dispatch_mgr.register_interceptor(doc)


def unregister_url_interceptor(*args):
    doc = CalcDoc.from_current_doc()
    dispatch_mgr.unregister_interceptor(doc)
