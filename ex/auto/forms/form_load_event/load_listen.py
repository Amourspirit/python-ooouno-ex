from __future__ import annotations
from pathlib import Path
from typing import Any, cast, TYPE_CHECKING

import uno
from com.sun.star.document import DocumentEvent
from com.sun.star.sdb.application import XDatabaseDocumentUI
from com.sun.star.frame import XFrameLoader

from ooo.dyn.sdb.application.database_object import DatabaseObject

from ooodev.loader import Lo
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
            doc_ui = Lo.qi(
                XDatabaseDocumentUI, controller.getModel().getCurrentController(), True
            )

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
        self._fn_on_document_event_occurred = self.on_document_event_occurred
        self._fn_on_notify = self.on_notify_event

        self._doc_event_listener = DocumentEventListener()
        self._doc_event_listener.on(
            "documentEventOccured", self._fn_on_document_event_occurred
        )
        self._doc_event_listener.on("notifyEvent", self._fn_on_notify)

    def _add_listeners_form(self) -> None:
        self._fn_on_loaded = self.on_loaded
        self._fn_on_loading = self.on_loading
        self._fn_on_unloaded = self.on_unloaded
        self._fn_on_unloading = self.on_unloading

        self._form_load_listener = FormLoadListener()
        self._form_load_listener.on("loaded", self._fn_on_loaded)
        self._form_load_listener.on("loading", self._fn_on_loading)
        self._form_load_listener.on("unloaded", self._fn_on_unloaded)
        self._form_load_listener.on("unloading", self._fn_on_unloading)

    def _add_listeners_frame(self) -> None:
        self._fn_on_frame_load_canceled = self.on_frame_load_canceled
        self._fn_on_frame_load_finished = self.on_frame_load_finished

        self._frame_load_listener = FrameLoadListener()
        self._frame_load_listener.on("loadCancelled", self._fn_on_frame_load_canceled)
        self._frame_load_listener.on("loadFinished", self._fn_on_frame_load_finished)

    def _set_internal_events(self):
        self._fn_on_form_loading = self.on_component_loading
        self._fn_on_form_loaded = self.on_component_loaded

        self._events.on("component_loading", self._fn_on_form_loading)
        self._events.on("component_loaded", self._fn_on_form_loaded)

    def on_document_event_occurred(self, source: Any, event_args: EventArgs) -> None:
        event = cast(DocumentEvent, event_args.event_data)
        print("Document Event Occurred:")
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
        print("On Frame Load Finished")
        # print(event.Source)
        print(event)
