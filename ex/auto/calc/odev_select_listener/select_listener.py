from __future__ import annotations
from typing import TYPE_CHECKING, Any, cast

import uno
from com.sun.star.frame import XController

from ooodev.adapter.awt.top_window_listener import TopWindowListener, EventArgs
from ooodev.adapter.view.selection_change_listener import SelectionChangeListener, GenericArgs
from ooodev.office.calc import Calc
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.utils.type_var import PathOrStr

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

        # close down when window closes
        self._twl = TopWindowListener(trigger_args=GenericArgs(listener=self))
        self._twl.on("windowClosing", SelectionListener.on_window_closing)

    def _get_cell_float(self, addr: CellAddress) -> float | None:
        # get_val returns floats as float intance
        obj = Calc.get_val(sheet=self.sheet, addr=addr)
        if isinstance(obj, float):
            return obj
        return None

    def _attach_listener(self) -> None:
        # pass GenericArgs with listener arg of self.
        # this will allow for this instance to be passed to events.
        # pass doc to constructor, this will allow listener to be automatically attached to document.
        self._s_listener = SelectionChangeListener(trigger_args=GenericArgs(listener=self), doc=self._doc)
        self._s_listener.on("selectionChanged", SelectionListener.on_selection_changed)
        self._s_listener.on("disposing", SelectionListener.on_disposing)

    @staticmethod
    def on_window_closing(source: Any, event_args: EventArgs, *args, **kwargs) -> None:
        print("Closing")
        try:
            listener = cast(SelectionListener, kwargs.get("listener", None))
            if listener:
                Lo.close_doc(listener._doc)
                Lo.close_office()
                listener.closed = True
        except Exception as e:
            print(f"  {e}")

    @staticmethod
    def on_selection_changed(source: Any, event_args: EventArgs, *args, **kwargs) -> None:
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
        sl = cast(SelectionListener, kwargs["listener"])
        addr = Calc.get_selected_cell_addr(sl._doc)
        if addr is None:
            return
        try:
            # better to wrap in try block.
            # otherwise errors crahses office
            if not Calc.is_equal_addresses(addr, sl.curr_addr):
                flt = sl._get_cell_float(sl.curr_addr)
                if flt is not None:
                    if sl.curr_val is None:  # so previously stored value was null
                        print(f"{Calc.get_cell_str(sl.curr_addr)} new value: {flt:.2f}")
                    else:
                        if sl.curr_val != flt:
                            print(f"{Calc.get_cell_str(sl.curr_addr)} has changed from {sl.curr_val:.2f} to {flt:.2f}")

            # update current address and value
            sl.curr_addr = addr
            sl.curr_val = sl._get_cell_float(addr)
            if sl.curr_val is not None:
                print(f"{Calc.get_cell_str(sl.curr_addr)} value: {sl.curr_val}")
        except Exception as e:
            print(e)

    @staticmethod
    def on_disposing(source: Any, event_args: EventArgs, *args, **kwargs) -> None:
        print("Disposing")
