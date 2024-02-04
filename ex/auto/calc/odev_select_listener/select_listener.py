from __future__ import annotations
from typing import TYPE_CHECKING, Any, cast

import uno  # noqa: F401
from com.sun.star.frame import XController

from ooo.dyn.table.cell_address import CellAddress

from ooodev.adapter.awt.top_window_events import TopWindowEvents
from ooodev.adapter.view.selection_change_events import SelectionChangeEvents
from ooodev.events.args.event_args import EventArgs
from ooodev.calc import Calc
from ooodev.calc import CalcDoc
from ooodev.loader import Lo


if TYPE_CHECKING:
    from com.sun.star.lang import EventObject


class SelectionListener:
    def __init__(self) -> None:
        super().__init__()
        self.closed = False
        loader = Lo.load_office(Lo.ConnectSocket())
        self._doc = CalcDoc.create_doc(loader=loader, visible=True)

        self._sheet = self._doc.sheets[0]

        self._curr_addr = self._doc.get_selected_cell_addr()
        self._curr_val = self._get_cell_float(self._curr_addr)  # may be None

        self._attach_listener()

        # insert some data
        self._sheet.set_col(
            values=("Smith", 42, 58.9, -66.5, 43.4, 44.5, 45.3),
            cell_name="A1",
        )

    def _get_cell_float(self, addr: CellAddress) -> float | None:
        # get_val returns floats as float instance
        obj = self._sheet.get_val(addr=addr)
        if isinstance(obj, float):
            return obj
        return None

    def _attach_listener(self) -> None:
        # Event handlers are defined as methods on the class.
        # However class methods are not callable by the event system.
        # The solution is to assign the method to class fields and use them to add the event callbacks.
        self._fn_on_window_closing = self.on_window_closing
        self._on_selection_changed = self.on_selection_changed
        self._on_disposing = self.on_disposing

        # close down when window closes
        self._twe = TopWindowEvents(add_window_listener=True)
        self._twe.add_event_window_closing(self._fn_on_window_closing)

        # pass doc to constructor, this will allow listener events to be automatically attached to document.
        self._sel_events = SelectionChangeEvents(doc=self._doc.component)
        self._sel_events.add_event_selection_changed(self._on_selection_changed)
        self._sel_events.add_event_selection_change_events_disposing(self._on_disposing)

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

    def on_selection_changed(
        self, source: Any, event_args: EventArgs, *args, **kwargs
    ) -> None:
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

        addr = self._doc.get_selected_cell_addr()
        if addr is None:
            return
        try:
            # better to wrap in try block.
            # otherwise errors crashes office
            if not Calc.is_equal_addresses(addr, self._curr_addr):
                flt = self._get_cell_float(self._curr_addr)
                if flt is not None:
                    if self._curr_val is None:  # so previously stored value was null
                        print(
                            f"{Calc.get_cell_str(self._curr_addr)} new value: {flt:.2f}"
                        )
                    else:
                        if self._curr_val != flt:
                            print(
                                f"{Calc.get_cell_str(self._curr_addr)} has changed from {self._curr_val:.2f} to {flt:.2f}"
                            )

            # update current address and value
            self._curr_addr = addr
            self._curr_val = self._get_cell_float(addr)
            if self._curr_val is not None:
                print(f"{Calc.get_cell_str(self._curr_addr)} value: {self._curr_val}")
        except Exception as e:
            print(e)

    def on_disposing(self, source: Any, event_args: EventArgs, *args, **kwargs) -> None:
        print("Disposing")
