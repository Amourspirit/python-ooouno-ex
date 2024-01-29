from __future__ import annotations
from typing import TYPE_CHECKING, Any

import uno
import unohelper
from com.sun.star.util import XModifyListener
from com.sun.star.sheet import XSpreadsheetDocument
from com.sun.star.util import XModifyBroadcaster

from ooodev.utils.lo import Lo
from ooodev.calc import Calc
from ooodev.calc import CalcDoc
from ooodev.utils.type_var import PathOrStr
from ooodev.utils.file_io import FileIO

# from ooodev.listeners.x_top_window_adapter import XTopWindowAdapter
from ooodev.adapter.awt.top_window_listener import TopWindowListener, EventArgs


if TYPE_CHECKING:
    from com.sun.star.lang import EventObject


class ModifyListener(unohelper.Base, XModifyListener):
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
        self._doc = CalcDoc.create_doc(loader, visible=True)

        self._sheet = self._doc.sheets[0]

        # insert some data
        self._sheet.set_col(
            cell_name="A1",
            values=("Smith", 42, 58.9, -66.5, 43.4, 44.5, 45.3),
        )

        mb = self._doc.qi(XModifyBroadcaster, True)
        mb.addModifyListener(self)

        # Event handlers are defined as methods on the class.
        # However class methods are not callable by the event system.
        # The solution is to create a function that calls the class method and pass that function to the event system.
        # Also the function must be a member of the class so that it is not garbage collected.

        def _on_window_closing(
            source: Any, event_args: EventArgs, *args, **kwargs
        ) -> None:
            self.on_window_closing(source, event_args, *args, **kwargs)

        self._fn_on_window_closing = _on_window_closing

        # close down when window closes
        self._twl = TopWindowListener()
        self._twl.on("windowClosing", _on_window_closing)

    def on_window_closing(
        self, source: Any, event_args: EventArgs, *args, **kwargs
    ) -> None:
        print("Closing")
        try:
            self._doc.close_doc()
            Lo.close_office()
            self.closed = True
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
        print(f"  {Calc.get_cell_str(addr=addr)} = {self._sheet.get_val(addr=addr)}")

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
