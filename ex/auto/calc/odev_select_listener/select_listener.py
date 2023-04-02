from __future__ import annotations
from typing import TYPE_CHECKING, Any, cast

import uno
from com.sun.star.frame import XController

from ooodev.adapter.awt.top_window_listener import TopWindowListener, EventArgs
from ooodev.adapter.view.selection_change_listener import SelectionChangeListener
from ooodev.office.calc import Calc
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo

from ooo.dyn.table.cell_address import CellAddress

if TYPE_CHECKING:
    from com.sun.star.lang import EventObject


class SelectionListener:
    def __init__(self) -> None:
        super().__init__()
        self.closed = False
        loader = Lo.load_office(Lo.ConnectSocket())
        self._doc = Calc.create_doc(loader)

        GUI.set_visible(is_visible=True, odoc=self._doc)
        self.sheet = Calc.get_sheet(doc=self._doc, index=0)

        self.curr_addr = Calc.get_selected_cell_addr(self._doc)
        self.curr_val = self._get_cell_float(self.curr_addr)  # may be None

        self._attach_listener()

        # insert some data
        Calc.set_col(sheet=self.sheet, cell_name="A1", values=("Smith", 42, 58.9, -66.5, 43.4, 44.5, 45.3))

    def _get_cell_float(self, addr: CellAddress) -> float | None:
        # get_val returns floats as float intance
        obj = Calc.get_val(sheet=self.sheet, addr=addr)
        if isinstance(obj, float):
            return obj
        return None

    def _attach_listener(self) -> None:

        # Event handlers are defined as methods on the class.
        # However class methods are not callable by the event system.
        # The solution is to create a function that calls the class method and pass that function to the event system.
        # Also the function must be a member of the class so that it is not garbage collected.

        def _on_window_closing(source: Any, event_args: EventArgs, *args, **kwargs) -> None:
            self.on_window_closing(source, event_args, *args, **kwargs)

        def _on_selection_changed(source: Any, event_args: EventArgs, *args, **kwargs) -> None:
            self.on_selection_changed(source, event_args, *args, **kwargs)

        def _on_disposing(source: Any, event_args: EventArgs, *args, **kwargs) -> None:
            self.on_disposing(source, event_args, *args, **kwargs)

        self._fn_on_window_closing = _on_window_closing
        self._on_selection_changed = _on_selection_changed
        self._on_disposing = _on_disposing

        # close down when window closes
        self._twl = TopWindowListener()
        self._twl.on("windowClosing", _on_window_closing)

        # pass doc to constructor, this will allow listener to be automatically attached to document.
        self._s_listener = SelectionChangeListener(doc=self._doc)
        self._s_listener.on("selectionChanged", _on_selection_changed)
        self._s_listener.on("disposing", _on_disposing)

    def on_window_closing(self, source: Any, event_args: EventArgs, *args, **kwargs) -> None:
        print("Closing")
        try:
            Lo.close_doc(self._doc)
            Lo.close_office()
            self.closed = True
        except Exception as e:
            print(f"  {e}")

    def on_selection_changed(self, source: Any, event_args: EventArgs, *args, **kwargs) -> None:
        # is fired four times for every click, and twice for shift arrow keys (?)
        # Once for arrow keys.
        # - Report when the selection changes by printing the name of the
        # previously selected cell and the new one;
        # - Report if cell just left has a new or changed numerical value;
        # - Report if cell just entered has a numerical value.
        event = cast("EventObject", event_args.event_data)
        ctrl = Lo.qi(XController, event.Source)
        if ctrl is None:
            print("No ctrl for event source")
            return

        addr = Calc.get_selected_cell_addr(self._doc)
        if addr is None:
            return
        try:
            # better to wrap in try block.
            # otherwise errors crahses office
            if not Calc.is_equal_addresses(addr, self.curr_addr):
                flt = self._get_cell_float(self.curr_addr)
                if flt is not None:
                    if self.curr_val is None:  # so previously stored value was null
                        print(f"{Calc.get_cell_str(self.curr_addr)} new value: {flt:.2f}")
                    else:
                        if self.curr_val != flt:
                            print(
                                f"{Calc.get_cell_str(self.curr_addr)} has changed from {self.curr_val:.2f} to {flt:.2f}"
                            )

            # update current address and value
            self.curr_addr = addr
            self.curr_val = self._get_cell_float(addr)
            if self.curr_val is not None:
                print(f"{Calc.get_cell_str(self.curr_addr)} value: {self.curr_val}")
        except Exception as e:
            print(e)

    def on_disposing(self, source: Any, event_args: EventArgs, *args, **kwargs) -> None:
        print("Disposing")
