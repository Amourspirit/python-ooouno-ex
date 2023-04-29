from __future__ import annotations
from pathlib import Path
from typing import Any, cast, TYPE_CHECKING
import uno
from com.sun.star.document import DocumentEvent

from ooodev.utils.lo import Lo
from ooodev.utils.gui import GUI
from ooodev.events.lo_events import Events
from ooodev.events.args.event_args import EventArgs
from ooodev.events.args.cancel_event_args import CancelEventArgs
from document_event_listener import DocumentEventListener

if TYPE_CHECKING:
    # only need types in design time and not at run time.
    from com.sun.star.lang import EventObject


class LoadListen:
    def __init__(self, db_path: Path) -> None:
        super().__init__()

        self._db_fnm = db_path

        self._events = Events(source=self)

        loader = Lo.load_office(Lo.ConnectSocket(), opt=Lo.Options(verbose=True))
        try:
            doc = Lo.open_doc(self._db_fnm, loader)
            assert doc is not None
            GUI.set_visible(doc)

            self._set_internal_events()
            self._add_listeners()

            oForm = doc.getFormDocuments()
            cargs = CancelEventArgs(source=self)
            self._events.trigger("form_loading", cargs)
            if not cargs.cancel:
                comp = oForm.loadComponentFromURL("starter", "_blank", 0, ())
                comp.addEventListener(self._doc_event_listener)
                eargs = EventArgs.from_args(cargs)
                eargs.event_data = comp
                self._events.trigger("form_loaded", eargs)

            Lo.close_doc(doc=doc)
        finally:
            Lo.close_office()

    def _add_listeners(self) -> None:
        def _on_document_event_occured(source: Any, event_args: EventArgs) -> None:
            self.on_document_event_occured(source, event_args)
        
        def _on_notify(source: Any, event_args: EventArgs) -> None:
            self.on_notify_event(source, event_args)

        self._fn_on_document_event_occured = _on_document_event_occured
        self._fn_on_notify = _on_notify

        self._doc_event_listener = DocumentEventListener()
        self._doc_event_listener.on("documentEventOccured", _on_document_event_occured)
        self._doc_event_listener.on("notifyEvent", _on_notify)

    def _set_internal_events(self):
        def _on_form_loading(source: Any, event_args: CancelEventArgs) -> None:
            self.on_form_loading(source=source, event_args=event_args)
        
        def _on_form_loaded(source: Any, event_args: EventArgs) -> None:
            self.on_form_loaded(source=source, event_args=event_args)
        
        self._fn_on_form_loading = _on_form_loading
        self._fn_on_form_loaded = _on_form_loaded
        
        self._events.on("form_loading", _on_form_loading)
        self._events.on("form_loaded", _on_form_loaded)

    def on_document_event_occured(self, source: Any, event_args: EventArgs) -> None:
        event = cast(DocumentEvent, event_args.event_data)
        print("Document Event Occured:")
        # print(event.Source)
        print(event.EventName)
    
    def on_notify_event(self, source: Any, event_args: EventArgs) -> None:
        event = cast("EventObject", event_args.event_data)
        print("Notify Event:")
        # print(event.Source)
        print(event.EventName)
    
    def on_form_loading(self, source: Any, event_args: CancelEventArgs) -> None:
        pass
    
    def on_form_loaded(self, source: Any, event_args: EventArgs) -> None:
        form = event_args.event_data.DrawPage.Forms.getByIndex(0)
        print(f'Form with name "{form.Name}" Loaded.')
