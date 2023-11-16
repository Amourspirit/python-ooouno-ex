from __future__ import annotations
from typing import Any, TYPE_CHECKING, cast
from pathlib import Path
import uno  # pylint: disable=unused-import

from ooo.dyn.awt.push_button_type import PushButtonType
from ooo.dyn.awt.pos_size import PosSize

from ooodev.dialog.msgbox import MsgBox, MessageBoxResultsEnum, MessageBoxType
from ooodev.dialog import Dialogs, BorderKind
from ooodev.events.args.event_args import EventArgs
from ooodev.office.calc import Calc
from ooodev.utils.lo import Lo
from ooodev.utils.gui import GUI
from listbox import Listbox
from listbox_drop_down import ListboxDropDown
from listbox_multi_select import ListboxMultiSelect

if TYPE_CHECKING:
    from com.sun.star.awt import ActionEvent
    from com.sun.star.awt.tab import TabPageActivatedEvent
    from com.sun.star.sheet import XSpreadsheetDocument


class Tabs:
    # pylint: disable=unused-argument
    def __init__(self, doc: XSpreadsheetDocument) -> None:
        self._border_kind = BorderKind.BORDER_SIMPLE
        self._width = 600
        self._height = 500
        self._btn_width = 100
        self._btn_height = 30
        self._margin = 4
        self._box_height = 30
        self._title = "Listbox Examples"
        if self._border_kind != BorderKind.BORDER_3D:
            self._padding = 10
        else:
            self._padding = 14
        self._tab_count = 0
        self._doc = doc
        self._init_dialog()

    def _init_dialog(self) -> None:
        self._init_handlers()

        self._dialog = Dialogs.create_dialog(
            x=-1,
            y=-1,
            width=self._width,
            height=self._height,
            title=self._title,
        )
        # createPeer() must be call before inserting tabs
        Dialogs.create_dialog_peer(self._dialog)

        # tab offset will vary depending on border kind and Operating System
        self._tab_offset_vert = (self._margin * 3) + 30
        self._init_tab_control()
        self._active_page_page_id = 1

    def _init_tab_control(self) -> None:
        self._ctl_tab = Dialogs.insert_tab_control(
            dialog_ctrl=self._dialog,
            x=self._margin,
            y=self._margin,
            width=self._width - (self._margin * 2),
            height=self._height - (self._margin * 2) - self._btn_height - (self._padding * 2),
        )
        self._ctl_tab.add_event_tab_page_activated(self._fn_tab_activated)
        self._init_tab_list_box()
        self._init_tab_drop_down()
        self._init_tab_multi_select()
        self._init_buttons()

    def _init_tab_list_box(self) -> None:
        self._tab_count += 1
        self._tab_main = Dialogs.insert_tab_page(
            dialog_ctrl=self._dialog,
            tab_ctrl=self._ctl_tab,
            title="Listbox",
            tab_position=self._tab_count,
        )
        tab_sz = self._ctl_tab.view.getPosSize()
        self._listbox = Listbox(
            ctrl=self._tab_main.view,
            doc=self._doc,
            x=tab_sz.X + self._margin,
            y=tab_sz.Y + self._tab_offset_vert,
            height=tab_sz.Height - self._tab_offset_vert - self._margin,
            width=tab_sz.Width - (self._margin * 2),
            border_kind=self._border_kind,
        )

    def _init_tab_drop_down(self) -> None:
        self._tab_count += 1
        self._tab_oth = Dialogs.insert_tab_page(
            dialog_ctrl=self._dialog,
            tab_ctrl=self._ctl_tab,
            title="Drop Down",
            tab_position=self._tab_count,
        )
        tab_sz = self._ctl_tab.view.getPosSize()
        self._listbox_drop_down = ListboxDropDown(
            ctrl=self._tab_oth.view,
            doc=self._doc,
            x=tab_sz.X + self._margin,
            y=tab_sz.Y + self._tab_offset_vert,
            height=tab_sz.Height - self._tab_offset_vert - self._margin,
            width=tab_sz.Width - (self._margin * 2),
            border_kind=self._border_kind,
        )

    def _init_tab_multi_select(self) -> None:
        self._tab_count += 1
        self._tab_oth = Dialogs.insert_tab_page(
            dialog_ctrl=self._dialog,
            tab_ctrl=self._ctl_tab,
            title="Multi Select",
            tab_position=self._tab_count,
        )
        tab_sz = self._ctl_tab.view.getPosSize()
        self._listbox_multi_select = ListboxMultiSelect(
            ctrl=self._tab_oth.view,
            doc=self._doc,
            x=tab_sz.X + self._margin,
            y=tab_sz.Y + self._tab_offset_vert,
            height=tab_sz.Height - self._tab_offset_vert - self._margin,
            width=tab_sz.Width - (self._margin * 2),
            border_kind=self._border_kind,
        )

    def _init_buttons(self) -> None:
        """Add OK, Cancel and Info buttons to dialog control"""
        self._ctl_btn_cancel = Dialogs.insert_button(
            dialog_ctrl=self._dialog.control,
            label="Cancel",
            x=self._width - self._btn_width - self._margin,
            y=self._height - self._btn_height - self._padding,
            width=self._btn_width,
            height=self._btn_height,
            btn_type=PushButtonType.CANCEL,
        )
        sz = self._ctl_btn_cancel.view.getPosSize()
        self._ctl_btn_ok = Dialogs.insert_button(
            dialog_ctrl=self._dialog.control,
            label="OK",
            x=sz.X - sz.Width - self._margin,
            y=sz.Y,
            width=self._btn_width,
            height=self._btn_height,
            btn_type=PushButtonType.OK,
            DefaultButton=True,
        )

        self._ctl_btn_info = Dialogs.insert_button(
            dialog_ctrl=self._dialog.control,
            label="Info",
            x=self._margin,
            y=sz.Y,
            width=self._btn_width,
            height=self._btn_height,
        )
        self._ctl_btn_info.view.setActionCommand("INFO")
        self._ctl_btn_info.model.HelpText = "Show info for selected item."
        self._ctl_btn_info.add_event_action_performed(self._fn_button_action_preformed)

    # region Show Dialog
    def show(self) -> int:
        self._ctl_tab.active_tab_page_id = self._active_page_page_id
        window = Lo.get_frame().getContainerWindow()
        ps = window.getPosSize()
        x = round(ps.Width / 2 - self._width / 2)
        y = round(ps.Height / 2 - self._height / 2)
        self._dialog.set_pos_size(x, y, self._width, self._height, PosSize.POSSIZE)
        self._dialog.set_visible(True)
        result = self._dialog.execute()
        self._handle_results(result)
        self._dialog.dispose()
        return result

    # endregion Show Dialog

    # region Handle Results
    def _handle_results(self, result: int) -> None:
        """Display a message if the OK button has been clicked"""
        if result == MessageBoxResultsEnum.OK.value:
            msg_vals = [self._listbox.get_data_message()]
            msg_vals.append(self._listbox_drop_down.get_data_message())
            msg_vals.append(self._listbox_multi_select.get_data_message())
            msg = "\n".join(msg_vals)
            _ = MsgBox.msgbox(
                msg=msg,
                title="Selected Values",
                boxtype=MessageBoxType.INFOBOX,
            )

    # endregion Handle Results

    # region Event Handlers
    def _init_handlers(self) -> None:
        self._fn_tab_activated = self.on_tab_activated
        self._fn_button_action_preformed = self.on_button_action_preformed

    def on_tab_activated(self, src: Any, event: EventArgs, control_src: Any, *args, **kwargs) -> None:
        print("Tab Changed:", control_src.name)
        itm_event = cast("TabPageActivatedEvent", event.event_data)
        self._active_page_page_id = itm_event.TabPageID
        print("Active ID:", self._active_page_page_id)

    def on_button_action_preformed(self, src: Any, event: EventArgs, control_src: Any, *args, **kwargs) -> None:
        """Method that is fired when Info button is clicked."""
        itm_event = cast("ActionEvent", event.event_data)
        if itm_event.ActionCommand == "INFO":
            self._handle_results(MessageBoxResultsEnum.OK.value)

    # endregion Event Handlers

    # region Properties

    @property
    def doc(self) -> XSpreadsheetDocument:
        return self._doc

    # endregion Properties
