from __future__ import annotations
from typing import Any

import uno

from ooodev.adapter.awt.top_window_events import TopWindowEvents
from ooodev.adapter.util.modify_events import ModifyEvents
from ooodev.events.args.event_args import EventArgs
from ooodev.office.calc import Calc
from ooodev.utils.file_io import FileIO
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.utils.type_var import PathOrStr


class ModifyListenerAdapter:
    def __init__(self, out_fnm: PathOrStr) -> None:
        super().__init__()
        if out_fnm:
            out_file = FileIO.get_absolute_path(out_fnm)
            _ = FileIO.make_directory(out_file)
            self._out_fnm = out_file
        else:
            self._out_fnm = ""
        self.closed = False
        loader = Lo.load_office(Lo.ConnectPipe())
        self._doc = Calc.create_doc(loader)

        GUI.set_visible(visible=True, doc=self._doc)
        self._sheet = Calc.get_sheet(doc=self._doc, index=0)

        # insert some data
        Calc.set_col(sheet=self._sheet, cell_name="A1", values=("Smith", 42, 58.9, -66.5, 43.4, 44.5, 45.3))

        # Event handlers are defined as methods on the class.
        # However class methods are not callable by the event system.
        # The solution is to assign the method to class fields and use them to add the event callbacks.
        self._fn_on_window_closing = self.on_window_closing
        self._fn_on_modified = self.on_modified
        self._fn_on_disposing = self.on_disposing

        # pass doc to constructor, this will allow listener to be automatically attached to document.
        self._m_events = ModifyEvents(subscriber=self._doc)
        self._m_events.add_event_modified(self._fn_on_modified)
        self._m_events.add_event_modify_events_disposing(self._fn_on_disposing)

        # close down when window closes
        self._top_win_ev = TopWindowEvents(add_window_listener=True)
        self._top_win_ev.add_event_window_closing(self._fn_on_window_closing)

    def on_window_closing(self, source: Any, event_args: EventArgs, *args, **kwargs) -> None:
        print("Closing")
        try:
            Lo.close_doc(self._doc)
            Lo.close_office()
            self.closed = True
        except Exception as e:
            print(f"  {e}")

    def on_modified(self, source: Any, event_args: EventArgs, *args, **kwargs) -> None:
        print("Modified")
        try:
            # event = cast("EventObject", event_args.event_data)
            # doc = Lo.qi(XSpreadsheetDocument, event.Source, True)
            doc = self._doc
            addr = Calc.get_selected_cell_addr(doc)
            print(f"  {Calc.get_cell_str(addr=addr)} = {Calc.get_val(sheet=self._sheet, addr=addr)}")
        except Exception as e:
            print(e)

    def on_disposing(self, source: Any, event_args: EventArgs, *args, **kwargs) -> None:
        print("Disposing")
