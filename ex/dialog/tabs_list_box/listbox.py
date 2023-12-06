# region Imports
from __future__ import annotations
import uno
from typing import Any, cast, TYPE_CHECKING, Tuple

from ooodev.dialog import Dialogs, BorderKind
from ooodev.events.args.event_args import EventArgs
from ooodev.calc import CalcDoc
from ooodev.utils.data_type.range_obj import RangeObj

if TYPE_CHECKING:
    from com.sun.star.awt import ActionEvent
    from com.sun.star.awt import ItemEvent
    from ooodev.dialog.dl_control.ctl_list_box import CtlListBox
    from com.sun.star.awt import XControl


# endregion Imports


class Listbox:
    """Listbox example."""

    # pylint: disable=unused-argument
    # region Init
    def __init__(
        self,
        ctrl: XControl,
        doc: CalcDoc,
        x: int,
        y: int,
        width: int,
        height: int,
        border_kind: BorderKind,
    ) -> None:
        self._control = ctrl
        self._doc = doc
        self._x = x
        self._y = y
        self._width = width
        self._height = height
        self._border_kind = border_kind
        self._btn_width = 100
        self._btn_height = 30
        self._margin = 6
        self._vert_margin = 12
        self._box_height = 30
        if self._border_kind != BorderKind.BORDER_3D:
            self._padding = 10
        else:
            self._padding = 14
        self._row_index = -1
        self._selected_item = ""
        self._sheet = self._doc.get_active_sheet()
        self._init()

    def _init(self) -> None:
        self._init_handlers()
        self._init_label()
        self._init_listbox()
        self._set_list_data()

    def _init_handlers(self) -> None:
        """
        Add event handlers for when changes occur.

        Methods can not be assigned directly to control callbacks.
        This is a python thing. However, methods can be assigned to class
        variable an in turn those can be assigned to callbacks.

        Example:
            ``self._ctl_btn_info.add_event_action_performed(self.on_button_action_preformed)``
            This would not work!

            ``self._ctl_btn_info.add_event_action_performed(self._fn_button_action_preformed)``
            This will work.
        """

        self._fn_on_item_state_changed = self.on_item_state_changed
        self._fn_on_action_preformed = self.on_action_preformed

    def _init_label(self) -> None:
        """Add a fixed text label to the dialog control"""
        self._ctl_main_lbl = Dialogs.insert_label(
            dialog_ctrl=self._control,
            label=self.get_label_msg(),
            x=self._margin,
            y=self._padding,
            width=self._width - (self._margin * 2),
            height=self._box_height,
        )

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
            height=self._height - sz.Height - (self._margin * 2),
            multi_select=False,
            drop_down=False,
        )
        self._ctl_listbox.add_event_item_state_changed(self._fn_on_item_state_changed)
        self._ctl_listbox.add_event_action_performed(self._fn_on_action_preformed)

    # endregion Init

    def get_label_msg(self) -> str:
        return "This is a Listbox example."

    # region Data

    def get_selected_items(self) -> Tuple[int, ...]:
        """Get the items that are to be selected in the listbox at startup."""
        return (0,)

    def _set_list_data(self) -> None:
        """Set the data in the listbox from the first column of the spreadsheet."""
        # find the used range in the sheet
        rng = self._sheet.find_used_range()

        # we only want the first column and skip the first row because it is a header
        # 1 - rng.range_obj.get_start_col() will give us first column minus the first row.
        # rng.range_obj.get_start_col() -1 would give us first column minus the last row.
        rng_col1 = 1 - rng.range_obj.get_start_col()

        # get the data from the range, this will be 2 dimensional tuple.
        cell_range = self._sheet.get_range(range_obj=rng_col1)
        tbl = cell_range.get_array()

        # convert the tuple to a set. This will remove duplicates.
        data_set = set(val[0] for val in tbl)

        # convert the set into a list so we can sort it.
        data = list(data_set)
        data.sort()
        # add the sorted list to the listbox
        self._ctl_listbox.set_list_data(data)
        if self._ctl_listbox.item_count > 0:
            self._ctl_listbox.model.SelectedItems = self.get_selected_items()
            self._selected_item = self._ctl_listbox.view.getItem(0)

    # endregion Data

    # region get data message
    def get_data_message(self) -> str:
        """Display a message if the OK button has been clicked"""
        if self._selected_item:
            return f"List Selected Item: {self.selected_item}"
        else:
            return "List: Nothing was selected"

    # endregion get data message

    # region Event Handlers
    def on_item_state_changed(
        self, src: Any, event: EventArgs, control_src: CtlListBox, *args, **kwargs
    ) -> None:
        """This Event fires when the selected item changes in the Listbox."""
        itm_event = cast("ItemEvent", event.event_data)
        print("State Changed: ItemID:", itm_event.ItemId)
        print("State Changed: Selected:", itm_event.Selected)
        print("State Changed: Highlighted:", itm_event.Highlighted)

        if itm_event.Selected > 0:
            self._selected_item = control_src.view.getItem(itm_event.Selected)
        else:
            self._selected_item = ""
        print("State Changed: Selected_item:", self._selected_item)

    def on_action_preformed(
        self, src: Any, event: EventArgs, control_src: Any, *args, **kwargs
    ) -> None:
        itm_event = cast("ActionEvent", event.event_data)
        print("Action: ActionCommand:", itm_event.ActionCommand)

    # endregion Event Handlers

    # region Properties
    @property
    def x(self) -> int:
        return self._x

    @property
    def y(self) -> int:
        return self._y

    @property
    def width(self) -> int:
        return self._width

    @property
    def height(self) -> int:
        return self._height

    @property
    def border_kind(self) -> BorderKind:
        return self._border_kind

    @property
    def btn_width(self) -> int:
        return self._btn_width

    @property
    def btn_height(self) -> int:
        return self._btn_height

    @property
    def margin(self) -> int:
        return self._margin

    @property
    def vert_margin(self) -> int:
        return self._vert_margin

    @property
    def box_height(self) -> int:
        return self._box_height

    @property
    def padding(self) -> int:
        return self._padding

    @property
    def doc(self) -> CalcDoc:
        return self._doc

    @property
    def control(self) -> XControl:
        return self._control

    @property
    def control_listbox(self) -> CtlListBox:
        return self._ctl_listbox

    @property
    def selected_item(self) -> str:
        return self._selected_item

    # endregion Properties
