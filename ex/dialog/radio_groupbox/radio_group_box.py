# region Imports
from __future__ import annotations
import uno
from typing import Any, cast, TYPE_CHECKING
from ooo.dyn.awt.push_button_type import PushButtonType
from ooo.dyn.awt.pos_size import PosSize
from ooo.dyn.awt.font_descriptor import FontDescriptor
from ooodev.dialog.msgbox import MsgBox, MessageBoxResultsEnum, MessageBoxType

from ooodev.dialog import Dialogs, BorderKind
from ooodev.events.args.event_args import EventArgs
from ooodev.loader import Lo
from ooodev.utils.info import Info
from ooodev.utils.color import StandardColor

if TYPE_CHECKING:
    from com.sun.star.awt import ActionEvent
    from com.sun.star.awt import ItemEvent
    from ooodev.dialog.dl_control.ctl_radio_button import CtlRadioButton


# endregion Imports


class RadioGroupBox:
    # pylint: disable=unused-argument
    # region Init
    def __init__(self) -> None:
        self._border_kind = BorderKind.BORDER_SIMPLE
        self._width = 400
        self._height = 310
        self._btn_width = 100
        self._btn_height = 30
        self._margin = 6
        self._box_height = 30
        self._title = "Radio Controls"
        self._msg = "Group Box and Radio control example"
        if self._border_kind != BorderKind.BORDER_3D:
            self._padding = 10
        else:
            self._padding = 14
        self._current_tab_index = 1
        self._group1_opt: CtlRadioButton | None = None
        self._group2_opt: CtlRadioButton | None = None
        # get or set a font descriptor. This helps to keep the font consistent across different platforms.
        fd = Info.get_font_descriptor("Liberation Serif", "Regular")
        if fd is None:
            fd = FontDescriptor(
                CharacterWidth=100.0,
                Kerning=True,
                WordLineMode=False,
                Pitch=2,
                Weight=100,
            )
        fd.Height = 10
        self._fd = fd
        self._init_dialog()

    def _init_dialog(self) -> None:
        """Create dialog and add controls."""
        self._init_handlers()
        self._dialog = Dialogs.create_dialog(
            x=-1, y=-1, width=self._width, height=self._height, title=self._title
        )
        self._dialog.set_visible(False)
        Dialogs.create_dialog_peer(self._dialog.control)
        self._init_label()
        self._init_group_boxes()
        self._init_radio_controls()
        self._init_buttons()

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

        self._fn_on_group1_changed = self.on_group1_changed
        self._fn_on_group2_changed = self.on_group2_changed
        self._fn_button_action_preformed = self.on_button_action_preformed

    def _init_label(self) -> None:
        """Add a fixed text label to the dialog control"""
        self._ctl_main_lbl = Dialogs.insert_label(
            dialog_ctrl=self._dialog.control,
            label=self._msg,
            x=self._margin,
            y=self._padding,
            width=self._width - (self._padding * 2),
            height=self._box_height,
        )
        self._ctl_main_lbl.set_font_descriptor(self._fd)
        self._ctl_main_lbl.font_descriptor.weight = 150  # make bold

    def _init_group_boxes(self) -> None:
        sz_lbl = self._ctl_main_lbl.view.getPosSize()
        self._ctl_gb1 = Dialogs.insert_group_box(
            dialog_ctrl=self._dialog.control,
            x=self._margin,
            y=self._padding + sz_lbl.X + sz_lbl.Height,
            height=self._height
            - (sz_lbl.X + sz_lbl.Height)
            - self._btn_height
            - (self._padding * 3),
            width=round(self._width / 2) - (self._margin * 2),
            label="Group Box ONE",
        )
        self._ctl_gb1.set_font_descriptor(self._fd)
        self._ctl_gb1.font_descriptor.weight = 150  # make bold

        sz = self._ctl_gb1.view.getPosSize()
        self._ctl_gb2 = Dialogs.insert_group_box(
            dialog_ctrl=self._dialog.control,
            x=sz.X + sz.Width + (self._margin * 2),
            y=sz.Y,
            height=sz.Height,
            width=sz.Width,
            label="Group Box TWO",
        )
        self._ctl_gb2.set_font_descriptor(self._ctl_gb1.font_descriptor.component)

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
        self._ctl_btn_cancel.set_font_descriptor(self._fd)
        self._ctl_btn_cancel.text_color = StandardColor.BLACK
        self._ctl_btn_cancel.background_color = StandardColor.RED_LIGHT1

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
        self._ctl_btn_ok.set_font_descriptor(self._fd)
        self._ctl_btn_ok.text_color = StandardColor.BLACK
        self._ctl_btn_ok.background_color = StandardColor.GREEN_LIGHT1

        self._ctl_btn_info = Dialogs.insert_button(
            dialog_ctrl=self._dialog.control,
            label="Info",
            x=self._margin,
            y=sz.Y,
            width=self._btn_width,
            height=self._btn_height,
        )
        self._ctl_btn_info.set_font_descriptor(self._fd)
        self._ctl_btn_info.view.setActionCommand("INFO")
        self._ctl_btn_info.model.HelpText = "Show info for selected items."
        self._ctl_btn_info.add_event_action_performed(self._fn_button_action_preformed)

    def _init_radio_controls(self) -> None:
        """
        Inserts radio buttons into dialog.

        Inserting more then a single group of radio buttons into a dialog can be a bit buggy.
        This in part has to do with how LibreOffice handles the Tab Indexes for the radio controls.

        The easies solution is to set the group box tab index right before adding a set of radio controls.
        This way between radio control set a different control is assigned a tab index before the next
        set of radio controls is created and added.

        When using tab indexes the tab indexes must be contiguous for each group but the next group must not
        pick up from the same tab index that the last group finished on.

        Alternatively it seems radio control groups can also be added with out setting a tab index as well.
        In order for this to work a different control such as a group box must be added to the dialog between
        adding radio control groups.
        """
        sz_gb1 = self._ctl_gb1.view.getPosSize()
        self._rb1 = Dialogs.insert_radio_button(
            dialog_ctrl=self._dialog.control,
            label="Group 1 - Radio Button 1",
            x=sz_gb1.X + self._margin,
            y=sz_gb1.Y + self._box_height,
            width=sz_gb1.Width - (self._margin * 2),
            height=20,
        )
        self._rb1.set_font_descriptor(self._fd)
        self._rb1.text_color = StandardColor.get_random_color()
        self._group1_opt = self._rb1
        self._ctl_gb1.model.TabIndex = self._current_tab_index
        self._current_tab_index += 1

        self._rb1.model.State = 1
        self._rb1.tab_index = self._current_tab_index
        self._current_tab_index += 1
        extra = 8
        self._rb1.add_event_item_state_changed(self._fn_on_group1_changed)
        rb1_sz = self._rb1.view.getPosSize()
        for i in range(1, extra + 1):
            radio_btn = Dialogs.insert_radio_button(
                dialog_ctrl=self._dialog.control,
                label=f"Group 1 - Radio Button {i + 1}",
                x=rb1_sz.X,
                y=rb1_sz.Y + (rb1_sz.Height * i),
                width=rb1_sz.Width,
                height=rb1_sz.Height,
            )
            radio_btn.set_font_descriptor(self._rb1.font_descriptor.component)
            radio_btn.text_color = StandardColor.get_random_color()
            radio_btn.tab_index = self._current_tab_index
            radio_btn.model.State = 0
            self._current_tab_index += 1
            radio_btn.add_event_item_state_changed(self._fn_on_group1_changed)

        # simplest way to break contiguous tab order and have it work
        self._ctl_gb2.model.TabIndex = self._current_tab_index
        self._current_tab_index += 1

        sz_gb2 = self._ctl_gb2.view.getPosSize()
        self._rb2 = Dialogs.insert_radio_button(
            dialog_ctrl=self._dialog.control,
            label="Group 2 - Radio Button 1",
            x=sz_gb2.X + self._margin,
            y=sz_gb2.Y + self._box_height,
            width=sz_gb2.Width - (self._margin * 2),
            height=rb1_sz.Height,
            # name="Radio1",
        )
        self._rb2.set_font_descriptor(self._rb1.font_descriptor.component)
        self._rb2.text_color = StandardColor.get_random_color()
        self._rb2.model.State = 1
        self._rb2.tab_index = self._current_tab_index
        self._group2_opt = self._rb2
        self._current_tab_index += 1
        self._rb2.add_event_item_state_changed(self._fn_on_group2_changed)
        rb2_sz = self._rb2.view.getPosSize()
        for i in range(1, extra + 1):
            radio_btn = Dialogs.insert_radio_button(
                dialog_ctrl=self._dialog.control,
                label=f"Group 2 - Radio Button {i + 1}",
                x=rb2_sz.X,
                y=rb2_sz.Y + (rb2_sz.Height * i),
                width=rb2_sz.Width,
                height=rb2_sz.Height,
                # name=f"Radio{i + 1}",
            )
            radio_btn.set_font_descriptor(self._rb1.font_descriptor.component)
            radio_btn.text_color = StandardColor.get_random_color()
            radio_btn.model.State = 0
            radio_btn.tab_index = self._current_tab_index
            self._current_tab_index += 1
            radio_btn.add_event_item_state_changed(self._fn_on_group2_changed)

        self._current_tab_index += 1

    # endregion Init

    # region Handle Results
    def _handle_results(self, result: int) -> None:
        """Display a message if the OK button has been clicked"""
        if result == MessageBoxResultsEnum.OK.value:
            msg = ""
            if self._group1_opt:
                msg += f"Selected Group 1: {self._group1_opt.model.Label}"
            if self._group2_opt:
                if msg:
                    msg += f"\nSelected Group 2: {self._group2_opt.model.Label}"
                else:
                    msg += f"Selected Group 2: {self._group2_opt.model.Label}"
            if msg == "":
                # should not happen
                msg = "No Options are selected"

            _ = MsgBox.msgbox(
                msg=msg,
                title="Selected Values",
                boxtype=MessageBoxType.INFOBOX,
            )

    # endregion Handle Results

    # region Event Handlers
    def on_group1_changed(
        self, src: Any, event: EventArgs, control_src: CtlRadioButton, *args, **kwargs
    ) -> None:
        itm_event = cast("ItemEvent", event.event_data)
        print("Group 1 Item ID:", itm_event.ItemId)
        print("Group 1 Tab Index:", control_src.tab_index)
        print("Group 1 Tab Name:", control_src.model.Name)
        self._group1_opt = control_src

    def on_group2_changed(
        self, src: Any, event: EventArgs, control_src: CtlRadioButton, *args, **kwargs
    ) -> None:
        itm_event = cast("ItemEvent", event.event_data)
        print("Group 2 Item ID:", itm_event.ItemId)
        print("Group 2 Tab Index:", control_src.tab_index)
        print("Group 2 Tab Name:", control_src.model.Name)
        self._group2_opt = control_src

    def on_button_action_preformed(
        self, src: Any, event: EventArgs, control_src: Any, *args, **kwargs
    ) -> None:
        """Method that is fired when Info button is clicked."""
        itm_event = cast("ActionEvent", event.event_data)
        if itm_event.ActionCommand == "INFO":
            self._handle_results(MessageBoxResultsEnum.OK.value)

    # endregion Event Handlers

    # region Show Dialog
    def show(self) -> int:
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
