# region Imports
from __future__ import annotations
import uno
from typing import TYPE_CHECKING

from ooodev.dialog import Dialogs, BorderKind

from listbox import Listbox

if TYPE_CHECKING:
    from com.sun.star.sheet import XSpreadsheetDocument
    from com.sun.star.awt import XControl


# endregion Imports


class ListboxDropDown(Listbox):
    """Drop down listbox example."""
    # region Init
    def __init__(
        self,
        ctrl: XControl,
        doc: XSpreadsheetDocument,
        x: int,
        y: int,
        width: int,
        height: int,
        border_kind: BorderKind,
    ) -> None:
        super().__init__(
            ctrl=ctrl,
            doc=doc,
            x=x,
            y=y,
            width=width,
            height=height,
            border_kind=border_kind,
        )

    # endregion Init

    # region Overrides
    def get_label_msg(self) -> str:
        return "This is a Drop down list example."

    def _init_listbox(self) -> None:
        sz = self._ctl_main_lbl.view.getPosSize()
        # multi_select must be false for drop_down to work.
        self._ctl_listbox = Dialogs.insert_list_box(
            dialog_ctrl=self._control,
            entries=(),
            x=sz.X,
            y=sz.Y + sz.Height + self._margin,
            width=sz.Width,
            line_count=20,
            height=self.box_height,
            multi_select=False,
            drop_down=True,
        )
        self._ctl_listbox.add_event_item_state_changed(self._fn_on_item_state_changed)
        self._ctl_listbox.add_event_action_performed(self._fn_on_action_preformed)

    def get_data_message(self) -> str:
        """Display a message if the OK button has been clicked"""
        if self._selected_item:
            return f"Drop Down List Selected Item: {self.selected_item}"
        else:
            return "Drop Down List: Nothing was selected"

    # endregion Overrides
