from __future__ import annotations
from typing import TYPE_CHECKING, Any, cast

import uno
from com.sun.star.sheet import XSpreadsheetDocument

from ooodev.utils.lo import Lo
from ooodev.utils.gui import GUI
from ooodev.office.calc import Calc
from ooodev.utils.type_var import PathOrStr
from ooodev.utils.file_io import FileIO
from ooodev.adapter.awt.top_window_listener import TopWindowListener, EventArgs
from ooodev.adapter.util.modify_listener import ModifyListener, GenericArgs

if TYPE_CHECKING:
    from com.sun.star.lang import EventObject


class ModifyListenerAdapter:
    def __init__(self, out_fnm: PathOrStr) -> None:
        super().__init__()
        if out_fnm:
            outf = FileIO.get_absolute_path(out_fnm)
            _ = FileIO.make_directory(outf)
            self._out_fnm = outf
        else:
            self._out_fnm = ""
        self.closed = False
        loader = Lo.load_office(Lo.ConnectPipe())
        self._doc = Calc.create_doc(loader)

        GUI.set_visible(is_visible=True, odoc=self._doc)
        self._sheet = Calc.get_sheet(doc=self._doc, index=0)

        # insert some data
        Calc.set_col(sheet=self._sheet, cell_name="A1", values=("Smith", 42, 58.9, -66.5, 43.4, 44.5, 45.3))

        # Event handlers are defined as methods on the class.
        # However class methods are not callable by the event system.
        # The solution is to create a function that calls the class method and pass that function to the event system.
        # Also the function must be a member of the class so that it is not garbage collected.

        def _on_window_closing(source: Any, event_args: EventArgs, *args, **kwargs) -> None:
            self.on_window_closing(source, event_args, *args, **kwargs)

        def _on_modified(source: Any, event_args: EventArgs, *args, **kwargs) -> None:
            self.on_modified(source, event_args, *args, **kwargs)

        def _on_disposing(source: Any, event_args: EventArgs, *args, **kwargs) -> None:
            self.on_disposing(source, event_args, *args, **kwargs)

        self._fn_on_window_closing = _on_window_closing
        self._fn_on_modified = _on_modified
        self._fn_on_disposing = _on_disposing

        # pass GenericArgs with listener arg of self.
        # this will allow for this instance to be passed to events.
        # pass doc to constructor, this will allow listener to be automatically attached to document.
        self._m_listener = ModifyListener(doc=self._doc)
        self._m_listener.on("modified", _on_modified)
        self._m_listener.on("disposing", _on_disposing)

        # close down when window closes
        self._twl = TopWindowListener()
        self._twl.on("windowClosing", _on_window_closing)

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
            event = cast("EventObject", event_args.event_data)
            doc = Lo.qi(XSpreadsheetDocument, event.Source, True)
            addr = Calc.get_selected_cell_addr(doc)
            print(f"  {Calc.get_cell_str(addr=addr)} = {Calc.get_val(sheet=self._sheet, addr=addr)}")
        except Exception as e:
            print(e)

    def on_disposing(self, source: Any, event_args: EventArgs, *args, **kwargs) -> None:
        print("Disposing")
