from __future__ import annotations
from typing import TYPE_CHECKING, Any, cast

import uno
import unohelper
from com.sun.star.util import XModifyListener
from com.sun.star.sheet import XSpreadsheetDocument
from com.sun.star.util import XModifyBroadcaster

from ooodev.utils.lo import Lo
from ooodev.utils.gui import GUI
from ooodev.office.calc import Calc
from ooodev.utils.type_var import PathOrStr
from ooodev.utils.file_io import FileIO

# from ooodev.listeners.x_top_window_adapter import XTopWindowAdapter
from ooodev.adapter.awt.top_window_listener import TopWindowListener, EventArgs, GenericArgs


if TYPE_CHECKING:
    from com.sun.star.lang import EventObject

class ModifyListener(unohelper.Base, XModifyListener):
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

        mb = Lo.qi(XModifyBroadcaster, self._doc, True)
        mb.addModifyListener(self)

        # close down when window closes
        self._twl = TopWindowListener(trigger_args=GenericArgs(listener=self))
        self._twl.on("windowClosing", ModifyListener.on_window_closing)

    @staticmethod
    def on_window_closing(source: Any, event_args: EventArgs, *args, **kwargs) -> None:
        print("Closing")
        try:
            ml = cast(ModifyListener, kwargs.get("listener", None))
            if ml:
                Lo.close_doc(ml._doc)
                Lo.close_office()
                ml.closed = True
        except Exception as e:
            print(f"  {e}")

    def modified(self, event: EventObject) -> None:
        """
        is called when something changes in the object.

        Due to such an event, it may be necessary to update views or controllers.

        The source of the event may be the content of the object to which the listener
        is registered.
        """
        print("Modified")
        doc = Lo.qi(XSpreadsheetDocument, event.Source, True)
        addr = Calc.get_selected_cell_addr(doc)
        print(f"  {Calc.get_cell_str(addr=addr)} = {Calc.get_val(sheet=self._sheet, addr=addr)}")

    def disposing(self, event: EventObject) -> None:
        """
        gets called when the broadcaster is about to be disposed.

        All listeners and all other objects, which reference the broadcaster
        should release the reference to the source. No method should be invoked
        anymore on this object ( including XComponent.removeEventListener() ).

        This method is called for every listener registration of derived listener
        interfaced, not only for registrations at XComponent.
        """
        print("Disposing")