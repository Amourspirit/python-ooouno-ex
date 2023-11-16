# region Imports
from __future__ import annotations
import uno
from typing import Any, cast, TYPE_CHECKING, Tuple

from ooodev.dialog import BorderKind
from ooodev.events.args.event_args import EventArgs

from listbox import Listbox

if TYPE_CHECKING:
    from com.sun.star.awt import ItemEvent
    from com.sun.star.sheet import XSpreadsheetDocument
    from ooodev.dialog.dl_control.ctl_list_box import CtlListBox
    from com.sun.star.awt import XControl


# endregion Imports


class ListboxMultiSelect(Listbox):
    """Multi-selection listbox example."""
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
        self._selected_items = (0,)
        super().__init__(
            ctrl=ctrl,
            doc=doc,
            x=x,
            y=y,
            width=width,
            height=height,
            border_kind=border_kind,
        )
        self.control_listbox.multi_selection = True

    # endregion Init

    # region Overrides
    def get_selected_items(self) -> Tuple[int, ...]:
        """Get the items that are to be selected in the listbox at startup."""
        return self._selected_items

    def get_label_msg(self) -> str:
        return "This is a Multi-selection Listbox example."

    def get_data_message(self) -> str:
        """Display a message if the OK button has been clicked"""
        if self._selected_item:
            return f"Drop Down List Selected Item: {self.selected_item}"
        else:
            return "Drop Down List: Nothing was selected"

    def on_item_state_changed(self, src: Any, event: EventArgs, control_src: CtlListBox, *args, **kwargs) -> None:
        """This Event fires when the selected item changes in the Listbox."""
        itm_event = cast("ItemEvent", event.event_data)
        print("State Changed: ItemID:", itm_event.ItemId)
        print("State Changed: Selected:", itm_event.Selected)
        print("State Changed: Highlighted:", itm_event.Highlighted)
        self._selected_items = control_src.model.SelectedItems
        print("State Changed: Selected_items:", self._selected_items)

    def get_data_message(self) -> str:
        """Display a message if the OK button has been clicked"""
        if self._selected_items:
            items = [""]
            items += [self._ctl_listbox.view.getItem(i) for i in self._selected_items]
            return "\nList Multi-Selected Item: ".join(items).lstrip()
        else:
            return "List: Nothing was selected"

    # endregion Overrides
