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

        self._m_listener = ModifyListener(event_args=GenericArgs(listener=self), doc=self._doc,)
        self._m_listener.on("modified", ModifyListenerAdapter.on_modified)
        self._m_listener.on("disposing", ModifyListenerAdapter.on_disposing)

        # close down when window closes
        self._twl = TopWindowListener(event_args=GenericArgs(listener=self))
        self._twl.on("windowClosing", ModifyListenerAdapter.on_window_closing)

    @staticmethod
    def on_window_closing(source: Any, event_args: EventArgs, **kwargs) -> None:
        print("Closing")
        try:
            ml = cast(ModifyListenerAdapter, kwargs.get("listener", None))
            if ml:
                Lo.close_doc(ml._doc)
                Lo.close_office()
                ml.closed = True
        except Exception as e:
            print(f"  {e}")

    @staticmethod
    def on_modified(source: Any, event_args: EventArgs, **kwargs) -> None:
        print("Modified")
        try:
            event = cast("EventObject", event_args.event_data)
            ml = cast(ModifyListenerAdapter, kwargs["listener"])
            doc = Lo.qi(XSpreadsheetDocument, event.Source, True)
            addr = Calc.get_selected_cell_addr(doc)
            print(f"  {Calc.get_cell_str(addr=addr)} = {Calc.get_val(sheet=ml._sheet, addr=addr)}")
        except Exception as e:
            print(e)

    @staticmethod
    def on_disposing(source: Any, event_args: EventArgs, **kwargs) -> None:
        print("Disposing")
