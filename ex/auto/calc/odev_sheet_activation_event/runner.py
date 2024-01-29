from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno

from ooodev.utils.lo import Lo
from ooodev.calc import Calc, CalcDoc, ZoomKind
from ooodev.adapter.awt.top_window_events import TopWindowEvents
from ooodev.events.args.event_args import EventArgs

from ooodev.dialog.msgbox import MsgBox, MessageBoxType

if TYPE_CHECKING:
    from com.sun.star.sheet import ActivationEvent


class Runner:
    """Runner Class for the Sheet Activation Event Example."""

    def __init__(self) -> None:
        super().__init__()

        self.closed = False
        self._tab_index = 1
        self._init_callbacks()

        loader = Lo.load_office(Lo.ConnectSocket())
        try:
            self._doc = CalcDoc.create_doc(loader=loader, visible=True)

            # Delay to let the doc become visible before zooming.
            Lo.delay(500)
            self._doc.zoom(ZoomKind.ZOOM_100_PERCENT)

            self._init_callbacks()
            # monitor the window closing event so we terminate the script
            self._top_win_ev = TopWindowEvents(add_window_listener=True)
            self._top_win_ev.add_event_window_closing(self._fn_on_window_closing)
            sheet = self._doc.sheets[0]
            sheet.name = "Sheet 1"
            sheet[
                "A1"
            ].value = (
                "This is Sheet 1. Select a different sheet to see the event in action."
            )
            # merge the cells just so the message looks nice
            sheet.get_range(range_name="A1:E1").merge_cells()
            _ = sheet.goto_cell("A1")
            self._add_sheets()

            # add the event listener that detects when the active sheet changes
            self._doc.current_controller.add_event_active_spreadsheet_changed(
                self._fn_on_active_sheet_changed
            )

        except Exception:
            Lo.close_office()
            raise

    def _init_callbacks(self) -> None:
        # Event handlers are defined as methods on the class.
        # However class methods are not callable by the event system.
        # The solution is to assign the method to class fields and use them to add the event callbacks.
        self._fn_on_window_closing = self.on_window_closing
        self._fn_on_active_sheet_changed = self.on_active_sheet_changed

    def _add_sheets(self) -> None:
        for i in range(2, 6):
            sheet = self._doc.sheets.insert_sheet(f"Sheet {i}")
            msg = f"This is Sheet {i}. Select a different sheet to see the event in action."
            sheet["A1"].value = msg
            sheet.get_range(range_name="A1:E1").merge_cells()

    def on_active_sheet_changed(
        self, source: Any, event_args: EventArgs, *args, **kwargs
    ) -> None:
        print("Active Sheet Changed")
        try:
            event = cast("ActivationEvent", event_args.event_data)
            # print("  Event:", event)
            # print("    ActiveSheet:", event.ActiveSheet)
            # print("    Source:", event.Source)

            # get the activate sheet as a CalcSheet object
            sheet = self._doc.sheets.get_sheet(event.ActiveSheet)
            print("    Active Sheet:", sheet.name)
            MsgBox.msgbox(
                f"The active sheet is now {sheet.name}.",
                "Sheet Activation Event",
                boxtype=MessageBoxType.INFOBOX,
            )
        except Exception as e:
            print(f"  {e}")

    # region Window Events
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

    # endregion Window Events
