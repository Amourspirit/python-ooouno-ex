# region Imports
from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno
from ooo.dyn.awt.font_descriptor import FontDescriptor
from ooo.dyn.awt.pos_size import PosSize
from ooo.dyn.awt.push_button_type import PushButtonType
from ooo.dyn.awt.rectangle import Rectangle
from ooo.dyn.style.vertical_alignment import VerticalAlignment
from ooodev.dialog import BorderKind
from ooodev.dialog.dl_control.ctl_fixed_text import CtlFixedText
from ooodev.dialog.msgbox import MessageBoxResultsEnum, MessageBoxType
from ooodev.events.args.event_args import EventArgs
from ooodev.gui.menu.popup_menu import PopupMenu
from ooodev.loader import Lo
from ooodev.units import UnitAppFontHeight
from ooodev.units import UnitAppFontX
from ooodev.units import UnitAppFontY
from .menu_data import get_popup_menu


if TYPE_CHECKING:
    from ooodev.dialog.dl_control.ctl_base import DialogControlBase
    from com.sun.star.awt import MenuEvent
    from com.sun.star.awt import FocusEvent

# endregion Imports


class MenuDialog:
    # pylint: disable=unused-argument
    # region Init
    def __init__(self) -> None:
        self._doc = Lo.current_doc
        self._border_kind = BorderKind.BORDER_3D
        self._width = 500
        self._height = 200
        self._btn_width = 100
        self._btn_height = 30
        self._margin = 6
        self._box_height = 30
        self._title = "Dialog Menu"
        self._msg = "Dialog Menu Example."
        if self._border_kind != BorderKind.BORDER_3D:
            self._padding = 10
        else:
            self._padding = 14
        self._current_tab_index = 1
        # create a font descriptor to set font properties for menu icon.
        self._mnu_lbl_fd = FontDescriptor(
            CharacterWidth=100.0,
            Kerning=True,
            WordLineMode=False,
            Pitch=0,
            Weight=100,
            Height=10,
            Width=10,
        )
        self._execute_cmds = {".uno:About", ".uno:HelpIndex"}
        self._main_menu = None
        self._init_dialog()

    def _init_dialog(self) -> None:
        """Create dialog and add controls."""
        self._init_handlers()
        self._dialog = self._doc.create_dialog(
            x=-1, y=-1, width=self._width, height=self._height, title=self._title
        )
        self._dialog.set_visible(False)
        self._dialog.create_peer()
        self._init_label()
        self._init_menu_label()
        self._init_event_text()
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

        self._fn_on_menu_lbl_mouse_entered = self.on_menu_lbl_mouse_entered
        self._fn_on_menu_lbl_mouse_click = self.on_menu_lbl_mouse_click
        self._fn_on_menu_select = self.on_menu_select

    def _init_label(self) -> None:
        """Add a fixed text label to the dialog control"""
        self._ctl_main_lbl = self._dialog.insert_label(
            label=self._msg,
            x=self._margin,
            y=self._padding * 2,
            width=self._width - (self._margin * 2),
            height=self._box_height,
        )

    def _init_menu_label(self) -> None:
        """Add a fixed text label to the dialog control"""
        lbl_item = self._dialog.insert_label(
            label="☰",
            x=0,
            y=0,
            width=14,
            height=14,
        )
        lbl_item.font_descriptor = self._mnu_lbl_fd
        lbl_item.add_event_mouse_entered(self._fn_on_menu_lbl_mouse_entered)
        lbl_item.add_event_mouse_pressed(self._fn_on_menu_lbl_mouse_click)

    def _init_buttons(self) -> None:
        """Add OK, Cancel and Info buttons to dialog control"""
        self._ctl_btn_cancel = self._dialog.insert_button(
            label="Cancel",
            x=self._width - self._btn_width - self._margin,
            y=self._height - self._btn_height - self._padding,
            width=self._btn_width,
            height=self._btn_height,
            btn_type=PushButtonType.CANCEL,
        )
        self._set_tab_index(self._ctl_btn_cancel)
        sz = self._ctl_btn_cancel.view.getPosSize()
        self._ctl_btn_ok = self._dialog.insert_button(
            label="OK",
            x=sz.X - sz.Width - self._margin,
            y=sz.Y,
            width=self._btn_width,
            height=self._btn_height,
            btn_type=PushButtonType.OK,
            DefaultButton=True,
        )
        self._set_tab_index(self._ctl_btn_ok)

    def _init_event_text(self) -> None:
        self._event_text = self._dialog.insert_text_field(
            text="",
            x=self._margin,
            y=50,
            width=self._width - (self._margin * 2),
            # height=sz.Height - self._btn_height - (self._vert_margin * 2),
            height=100,
            border=self._border_kind,
            VerticalAlign=VerticalAlignment.TOP,
            ReadOnly=True,
            MultiLine=True,
            AutoVScroll=True,
        )
        self._set_tab_index(self._event_text)

    # endregion Init

    def _write_line(self, text: str) -> None:
        self._event_text.write_line(text)

    # region Handle Results
    def _handle_results(self, result: int) -> None:
        """Display a message if the OK button has been clicked"""
        if result == MessageBoxResultsEnum.OK.value:
            _ = self._doc.msgbox(
                msg="All done",
                title="Finished",
                boxtype=MessageBoxType.INFOBOX,
            )

    # endregion Handle Results

    # region Event Handlers
    def _get_main_menu(self) -> PopupMenu:
        if self._main_menu is None:
            self._main_menu = get_popup_menu()
            self._main_menu.subscribe_all_item_selected(self._fn_on_menu_select)
        return self._main_menu

    def _display_popup(self, control_src: CtlFixedText) -> None:
        menu_x = UnitAppFontX(control_src.model.PositionX)
        menu_y = UnitAppFontY(control_src.model.PositionY) + UnitAppFontHeight(
            control_src.model.Height
        )
        rect = Rectangle(
            X=int(menu_x.get_value_px()),
            Y=int(menu_y.get_value_px()),
            Width=10,
            Height=10,
        )
        pm = self._get_main_menu()
        peer = self._dialog.get_peer()
        pm.execute(peer, rect, 0)

    def on_menu_lbl_mouse_click(
        self,
        src: Any,
        event: EventArgs,
        control_src: CtlFixedText,
        *args,
        **kwargs,
    ) -> None:
        # print(control_src)
        print("Menu Mouse clicked:", control_src.name)

    def on_menu_select(self, src: Any, event: EventArgs, menu: PopupMenu) -> None:
        print("Menu Selected")
        me = cast("MenuEvent", event.event_data)
        command = menu.get_command(me.MenuId)
        self._write_line(f"Menu Selected: {command}, Menu ID: {me.MenuId}")
        if command == ".uno:exit":
            self._dialog.end_execute()
            return
        if command == ".uno:exitok":
            self._dialog.end_dialog(MessageBoxResultsEnum.OK.value)
            return
        if command in self._execute_cmds and menu.is_dispatch_cmd(command):
            menu.execute_cmd(command)

    # region Menu Label Mouse Events
    def on_menu_lbl_mouse_entered(
        self,
        src: Any,
        event: EventArgs,
        control_src: CtlFixedText,
        *args,
        **kwargs,
    ) -> None:
        self._display_popup(control_src)

    # endregion Menu Label Mouse Events

    # endregion Event Handlers

    # region Show Dialog
    def show(self) -> int:
        # make sure the document is active.
        self._doc.activate()
        window = self._doc.get_frame().getContainerWindow()
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

    def _set_tab_index(self, ctl: DialogControlBase) -> None:
        ctl.tab_index = self._current_tab_index
        self._current_tab_index += 1
