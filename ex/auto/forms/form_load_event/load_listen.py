from __future__ import annotations
from pathlib import Path
from typing import Any, cast, TYPE_CHECKING

import uno
from com.sun.star.document import DocumentEvent
from com.sun.star.sdb.application import XDatabaseDocumentUI
from com.sun.star.frame import XFrameLoader

from ooo.dyn.sdb.application.database_object import DatabaseObject

from ooodev.utils.lo import Lo
from ooodev.utils.gui import GUI
from ooodev.events.lo_events import Events
from ooodev.events.args.event_args import EventArgs
from ooodev.events.args.cancel_event_args import CancelEventArgs
from document_event_listener import DocumentEventListener
from form_load_listener import FormLoadListener
from frame_load_listener import FrameLoadListener

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
            self._add_listeners_doc()
            # self._add_listeners_form()
            # self._add_listeners_frame()

            controller = GUI.get_current_controller(doc)
            doc_ui = Lo.qi(XDatabaseDocumentUI, controller.getModel().getCurrentController(), True)

            # must connect to use doc_ui.loadComponent()
            doc_ui.connect()

            cargs = CancelEventArgs(source=self)
            self._events.trigger("component_loading", cargs)
            if not cargs.cancel:
                # comp = component_loader.loadComponentFromURL("starter", "_blank", 0, ())
                comp = doc_ui.loadComponent(DatabaseObject.FORM, "starter", False)
                comp.addEventListener(self._doc_event_listener)
                eargs = EventArgs.from_args(cargs)
                eargs.event_data = comp
                self._events.trigger("component_loaded", eargs)

            Lo.close_doc(doc=doc)
        finally:
            Lo.close_office()

    def _add_listeners_doc(self) -> None:
        def _on_document_event_occured(source: Any, event_args: EventArgs) -> None:
            self.on_document_event_occured(source, event_args)

        def _on_notify(source: Any, event_args: EventArgs) -> None:
            self.on_notify_event(source, event_args)

        self._fn_on_document_event_occured = _on_document_event_occured
        self._fn_on_notify = _on_notify

        self._doc_event_listener = DocumentEventListener()
        self._doc_event_listener.on("documentEventOccured", _on_document_event_occured)
        self._doc_event_listener.on("notifyEvent", _on_notify)

    def _add_listeners_form(self) -> None:
        def _on_loaded(source: Any, event_args: EventArgs) -> None:
            self.on_loaded(source, event_args)

        def _on_loading(source: Any, event_args: EventArgs) -> None:
            self.on_loading(source, event_args)

        def _on_unloaded(source: Any, event_args: EventArgs) -> None:
            self.on_unloaded(source, event_args)

        def _on_unloading(source: Any, event_args: EventArgs) -> None:
            self.on_unloading(source, event_args)

        self._fn_on_loaded = _on_loaded
        self._fn_on_loading = _on_loading
        self._fn_on_unloaded = _on_unloaded
        self._fn_on_unloading = _on_unloading

        self._form_load_listener = FormLoadListener()
        self._form_load_listener.on("loaded", _on_loaded)
        self._form_load_listener.on("loading", _on_loading)
        self._form_load_listener.on("unloaded", _on_unloaded)
        self._form_load_listener.on("unloading", _on_unloading)

    def _add_listeners_frame(self) -> None:
        def _on_frame_load_canceled(source: Any, event_args: EventArgs) -> None:
            self.on_frame_load_canceled(source, event_args)

        def _on_frame_load_finished(source: Any, event_args: EventArgs) -> None:
            self.on_frame_load_finished(source, event_args)

        self._fn_on_frame_load_canceled = _on_frame_load_canceled
        self._fn_on_frame_load_finished = _on_frame_load_finished

        self._frame_load_listener = FrameLoadListener()
        self._frame_load_listener.on("loadCancelled", _on_frame_load_canceled)
        self._frame_load_listener.on("loadFinished", _on_frame_load_finished)

    def _set_internal_events(self):
        def _on_component_loading(source: Any, event_args: CancelEventArgs) -> None:
            self.on_component_loading(source=source, event_args=event_args)

        def _on_component_loaded(source: Any, event_args: EventArgs) -> None:
            self.on_component_loaded(source=source, event_args=event_args)

        self._fn_on_form_loading = _on_component_loading
        self._fn_on_form_loaded = _on_component_loaded

        self._events.on("component_loading", _on_component_loading)
        self._events.on("component_loaded", _on_component_loaded)

    def on_document_event_occured(self, source: Any, event_args: EventArgs) -> None:
        event = cast(DocumentEvent, event_args.event_data)
        print("Document Event Occured:")
        # print(event.Source)
        print(event.EventName)

    def on_notify_event(self, source: Any, event_args: EventArgs) -> None:
        event = cast("EventObject", event_args.event_data)
        print(f"Notify Event: {event.EventName}")

    def on_component_loading(self, source: Any, event_args: CancelEventArgs) -> None:
        pass

    def on_component_loaded(self, source: Any, event_args: EventArgs) -> None:
        form = event_args.event_data.DrawPage.Forms.getByIndex(0)
        print(f'Form with name "{form.Name}" Loaded.')

    def on_loaded(self, source: Any, event_args: EventArgs) -> None:
        event = cast(DocumentEvent, event_args.event_data)
        print("On Loaded")
        # print(event.Source)
        print(event.EventName)

    def on_loading(self, source: Any, event_args: EventArgs) -> None:
        event = cast(DocumentEvent, event_args.event_data)
        print("On Loading")
        # print(event.Source)
        print(event.EventName)

    def on_unloaded(self, source: Any, event_args: EventArgs) -> None:
        event = cast(DocumentEvent, event_args.event_data)
        print("On Unloaded")
        # print(event.Source)
        print(event.EventName)

    def on_unloading(self, source: Any, event_args: EventArgs) -> None:
        event = cast(DocumentEvent, event_args.event_data)
        print("On Unloading")
        # print(event.Source)
        print(event.EventName)

    def on_frame_load_canceled(self, source: Any, event_args: EventArgs) -> None:
        event = cast(XFrameLoader, event_args.event_data)
        print("On Frame Load Canceled")
        # print(event.Source)
        print(event)

    def on_frame_load_finished(self, source: Any, event_args: EventArgs) -> None:
        event = cast(XFrameLoader, event_args.event_data)
        print("On Frame Load Finsihsed")
        # print(event.Source)
        print(event)
