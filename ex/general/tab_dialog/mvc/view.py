# coding: utf-8
# region imports
from ooo.dyn.awt.uno_control_button_model import UnoControlButtonModel
from ooo.dyn.awt.tab.uno_control_tab_page_container_model import (
    UnoControlTabPageContainerModel,
)
from ooo.dyn.awt.tab.uno_control_tab_page_model import UnoControlTabPageModel
from ooo.dyn.awt.uno_control_fixed_text_model import UnoControlFixedTextModel
from ooo.dyn.awt.uno_control_list_box_model import UnoControlListBoxModel
from ooo.dyn.awt.uno_control_check_box_model import UnoControlCheckBoxModel
from ooo.dyn.awt.uno_control_group_box_model import UnoControlGroupBoxModel
from ooo.dyn.awt.uno_control_radio_button_model import UnoControlRadioButtonModel
from ooo.dyn.uno.runtime_exception import RuntimeException
from ooo.dyn.util.color import Color
from ooo.dyn.awt.push_button_type import PushButtonType
from typing import TYPE_CHECKING, Union
from .enums import TabEnum, TimeEnum, TextEnum, SyntaxEnum, VoidEnum
from .interface import IViewMultiSyntax, IControllerMultiSyntax
from ....ui.listeners.action_listener import ActionListener
from ....ui.listeners.property_change_listener import PropertyChangeListener
from ....ui.listeners.item_listener import ItemListener
from ....ui.listeners.tab_page_container_listener import TabPageContainerListener
from ....ui.dialog.builder.control_wrapper import ControlWrapper
from ....ui.dialog.builder.dialog_builder import DialogBuilder
from ....utils import color as ucolor

if TYPE_CHECKING:
    from ooo.csslo.awt import (
        UnoControlButton,
        UnoControlCheckBox,
        UnoControlContainerModel,
        UnoControlDialog,
        UnoControlFixedText,
        UnoControlGroupBox,
        UnoControlListBox,
        UnoControlRadioButton,
        ItemEvent,
        XControlContainer,
        XStyleSettings,
    )
    from ooo.csslo.awt.tab import (
        UnoControlTabPageContainer,
        TabPageActivatedEvent,
        UnoControlTabPage,
    )
    from ooo.lo.beans.property_change_event import PropertyChangeEvent
    from ooo.csslo.container import XNameContainer
    from ooo.csslo.lang import XInitialization, XMultiServiceFactory, XComponent

    UnoControlTabPageModelType = Union[
        UnoControlTabPageModel, XNameContainer, XInitialization, XMultiServiceFactory
    ]
    UnoControlTabPageContainerModelType = Union[
        UnoControlTabPageContainerModel,
        UnoControlContainerModel,
        XComponent,
        XControlContainer,
    ]
# endregion imports


class MultiSyntaxView(IViewMultiSyntax):
    # region Init
    def __init__(self) -> None:
        self._width = 320
        self._height = 360
        self._border_padding_x = 5
        self._padding_x = 5
        self._padding_y = 5
        self._btn_height = 26
        self._check_box_height = 20
        self._lbl_title_height = 42
        self._horz_sep = 8
        self._vert_sep = 6
        self._radio_btn_height = 20
        self._groupbox_offset_y = 18
        self._pg1_gb_width = 130
        self._gb_y_ofset = 18
        self._title = "Tabbed Dialog"
        self._name_btn_ok = "btnOk"
        self._name_btn_cancel = "btnCancel"
        self._builder: DialogBuilder = None
        self._refreshing = False

    # endregion Init

    # region Setup
    def setup(self, controller: IControllerMultiSyntax):
        try:
            self._controller = controller
            self._create()
            self._populate_list_syntax()
            self._populate_list_void()
            self.ctl_btn_ok.addActionListener(ActionListener(self._ok))
            self.ctl_btn_cancel.addActionListener(ActionListener(self._cancel))
            self.ctl_tab.addTabPageContainerListener(
                TabPageContainerListener(self._tab_changed)
            )
            self.ctl_lb_syntax.addItemListener(ItemListener(self._lb_syntax_change))
            self.ctl_lb_void.addItemListener(ItemListener(self._lb_void_change))

            self._m_chk_dpv.addPropertyChangeListener(
                "State", PropertyChangeListener(self._dpv_state_change)
            )
            self._m_opt_time_none.addPropertyChangeListener(
                "State", PropertyChangeListener(self._opt_time_state_change)
            )
            self._m_opt_time_past.addPropertyChangeListener(
                "State", PropertyChangeListener(self._opt_time_state_change)
            )
            self._m_opt_time_future.addPropertyChangeListener(
                "State", PropertyChangeListener(self._opt_time_state_change)
            )

            self._m_opt_pre_suf_none.addPropertyChangeListener(
                "State", PropertyChangeListener(self._opt_pre_suf_state_change)
            )
            self._m_opt_prefix.addPropertyChangeListener(
                "State", PropertyChangeListener(self._opt_pre_suf_state_change)
            )
            self._m_opt_suffix.addPropertyChangeListener(
                "State", PropertyChangeListener(self._opt_pre_suf_state_change)
            )
        except Exception as e:
            print(e)
            raise

    # endregion Setup

    # region Show
    def show(self) -> None:
        try:
            self._set_dialog_properties()
            # refresh must be called before self._builder.dialog.execute()
            self.refresh()

            # uncomment next line to turns off system theming of controls
            # self._builder.dialog.getPeer().setProperty("NativeWidgetLook", False)

            style_settings: 'XStyleSettings' = self._builder.dialog.StyleSettings
            # get the color of the dialog
            rgb_bg = ucolor.rgb.from_int(style_settings.DialogColor)

            if rgb_bg.is_dark():
                lb_gb = ucolor.lighten(rgb_bg, 5).to_int()
            else:
                lb_gb = ucolor.darken(rgb_bg, 5).to_int()
            self._m_lb_syntax.BackgroundColor = lb_gb
            self._m_lb_void_key.BackgroundColor = lb_gb
            self._builder.dialog.execute()
        except Exception as e:
            print(e)
            raise

    # endregion Show

    # region Populate list

    def _populate_list_syntax(self):
        self._m_lb_syntax.removeAllItems()
        lst = self._controller.controller_syntax.get_list_data()
        for i, itm in enumerate(lst):
            self._m_lb_syntax.insertItemText(i, itm)

    def _populate_list_void(self):
        self._m_lb_void_key.removeAllItems()
        lst = self._controller.controller_void.get_list_data()
        for i, itm in enumerate(lst):
            self._m_lb_void_key.insertItemText(i, itm)

    # endregion Populate list

    # region Create Dialog
    def _get_tab_pixels_width(self, pixels: int) -> int:
        # Pixel in a tab page.
        # not sure what units are being used in a tab page but here is what I got for pixels
        # for width in pixels for a control in a tabpage
        #   figure the pixel width desired and divide by 2.4
        #   figure the pixel height desired and divide by 2.133333333
        return round(pixels / 2.4)

    def _get_tab_pixels_height(self, pixels: int) -> int:
        return round(pixels / 2.133333333)

    def _create_tabs(self) -> None:
        tabs_model: "UnoControlTabPageContainerModelType" = (
            self._builder.create_control(UnoControlTabPageContainerModel)
        )
        tabs_model.Name = "tab_main"
        tabs_model.PositionX = self._border_padding_x
        tabs_model.PositionY = self._m_lbl_title_wrapper.bottom + self._padding_y
        tabs_model.Width = self._width - (self._border_padding_x * 2)
        tabs_model.Height = self._height - self._lbl_title_height - 40
        tabs_model.TabIndex = 3
        # tabs_model.BackgroundColor = int("4d4d4d", 16)
        # border, 0 no border, 1 3D border, 2 Simple
        tabs_model.Border = 2
        self._builder.add_control(tabs_model.Name, tabs_model)
        self._m_tabs_model = tabs_model
        self._create_tab_syntax()
        self._create_tab_void()
        self._create_tab_syntax_controls()
        self._create_tab_void_controls()

    def _create_tab_syntax(self) -> None:
        try:
            tab_syntax: "UnoControlTabPageModelType" = self._builder.create_control(
                UnoControlTabPageModel
            )

            tab_syntax.initialize((1,))
            tab_syntax.Title = "SYNTAX"
            self._m_tab_syntax = tab_syntax
        except RuntimeException as e:
            raise e

    def _create_tab_void(self) -> None:
        try:
            tab_void: "UnoControlTabPageModelType" = self._builder.create_control(
                UnoControlTabPageModel
            )
            tab_void.initialize((2,))
            tab_void.Title = "VOID"
            self._m_tab_void = tab_void
        except RuntimeException as e:
            raise e

    def _create_tab_syntax_controls(self) -> None:
        lbl_syntax_key: UnoControlFixedTextModel = self._m_tab_syntax.createInstance(
            UnoControlFixedTextModel.__ooo_full_ns__
        )
        lbl_syntax_key.PositionX = 2
        lbl_syntax_key.PositionY = 2
        lbl_syntax_key.Width = self._get_tab_pixels_width(150)
        lbl_syntax_key.Height = self._get_tab_pixels_height(self._btn_height)
        lbl_syntax_key.Label = "KEY"
        lbl_syntax_key.NoLabel = True
        lbl_syntax_key.Align = 0
        # lbl_syntax_key.BackgroundColor = int("ff3322", 16)
        lbl_syntax_key_wraper = ControlWrapper(
            "lblSyntaxKey",
            lbl_syntax_key,
            lbl_syntax_key.PositionX,
            lbl_syntax_key.PositionY,
        )
        self._m_tab_syntax.insertByName(lbl_syntax_key_wraper.name, lbl_syntax_key)
        self._m_lbl_syntax = lbl_syntax_key

        lb_syntax_key: UnoControlListBoxModel = self._m_tab_syntax.createInstance(
            UnoControlListBoxModel.__ooo_full_ns__
        )
        lb_syntax_key.Name = "lbSyntaxKey"
        lb_syntax_key.PositionX = lbl_syntax_key.PositionX
        lb_syntax_key.PositionY = lbl_syntax_key_wraper.bottom + self._padding_y
        lb_syntax_key.Dropdown = False
        lb_syntax_key.Border = 2
        lb_syntax_key.Tabstop = True
        lb_syntax_key.TabIndex = 0
        lb_syntax_key.Width = lbl_syntax_key_wraper.width
        lb_syntax_key.Height = self._get_tab_pixels_height(160)
        # lb_syntax_key.BackgroundColor = int("272727", 16)
        lb_syntax_key_wraper = ControlWrapper(
            lb_syntax_key.Name,
            lb_syntax_key,
            lb_syntax_key.PositionX,
            lb_syntax_key.PositionY,
        )
        self._m_tab_syntax.insertByName(lb_syntax_key.Name, lb_syntax_key)
        self._m_lb_syntax = lb_syntax_key

        chk_dpv: UnoControlCheckBoxModel = self._m_tab_syntax.createInstance(
            UnoControlCheckBoxModel.__ooo_full_ns__
        )
        chk_dpv.PositionX = lb_syntax_key_wraper.x
        chk_dpv.PositionY = lb_syntax_key_wraper.bottom + self._padding_y
        chk_dpv.Name = "chkDPV"
        chk_dpv.Height = self._get_tab_pixels_height(self._check_box_height)
        chk_dpv.Width = lbl_syntax_key_wraper.width
        chk_dpv.Label = "DPV"
        chk_dpv.TriState = False
        chk_dpv.State = 0
        chk_dpv.Tabstop = True
        chk_dpv.TabIndex = 9
        chk_dpv.Align = 0
        self._m_tab_syntax.insertByName(chk_dpv.Name, chk_dpv)
        self._m_chk_dpv = chk_dpv

        # To combine several option buttons in a group, you must position them one after another in the activation sequence
        # without any gaps (Model.TabIndex property, described as Order in the dialog editor).
        # If the activation sequence is interrupted by another control element, then Apache OpenOffice automatically
        # starts with a new control element group that can be activated regardless of the first group of control elements.
        # https://wiki.openoffice.org/wiki/Documentation/BASIC_Guide/Control_Elements
        gb_time: UnoControlGroupBoxModel = self._m_tab_syntax.createInstance(
            UnoControlGroupBoxModel.__ooo_full_ns__
        )
        gb_time.Name = "gbTime"
        gb_time.Width = self._get_tab_pixels_width(self._pg1_gb_width)
        gb_time.Height = self._get_tab_pixels_height(90)
        gb_time.PositionY = lb_syntax_key_wraper.y - self._get_tab_pixels_height(8)
        gb_time.PositionX = lb_syntax_key_wraper.right + self._padding_x
        gb_time.Label = "Time"
        gb_time.TabIndex = 1
        gb_time_wraper = ControlWrapper(
            gb_time.Name, gb_time, gb_time.PositionX, gb_time.PositionY
        )
        self._m_tab_syntax.insertByName(gb_time.Name, gb_time)
        self._m_gb_time = gb_time

        opt_time_none: UnoControlRadioButtonModel = self._m_tab_syntax.createInstance(
            UnoControlRadioButtonModel.__ooo_full_ns__
        )
        opt_time_none.Name = "optTimeNone"
        opt_time_none.PositionX = gb_time.PositionX + self._get_tab_pixels_width(
            self._horz_sep
        )
        opt_time_none.PositionY = gb_time.PositionY + self._get_tab_pixels_height(
            self._gb_y_ofset
        )

        opt_time_none.Width = gb_time.Width - (
            self._get_tab_pixels_width(self._horz_sep) * 2
        )
        opt_time_none.Height = self._get_tab_pixels_height(self._radio_btn_height)
        opt_time_none.Label = "None"
        opt_time_none.State = 0
        opt_time_none.Tabstop = True
        opt_time_none.TabIndex = 2
        opt_time_none_wrappper = ControlWrapper(
            opt_time_none.Name,
            opt_time_none,
            opt_time_none.PositionX,
            opt_time_none.PositionY,
        )
        self._m_tab_syntax.insertByName(opt_time_none.Name, opt_time_none)
        self._m_opt_time_none = opt_time_none

        opt_past: UnoControlRadioButtonModel = self._m_tab_syntax.createInstance(
            UnoControlRadioButtonModel.__ooo_full_ns__
        )
        opt_past.Name = "optPast"
        opt_past.PositionX = opt_time_none.PositionX
        opt_past.PositionY = opt_time_none_wrappper.bottom
        opt_past.Width = opt_time_none.Width
        opt_past.Height = opt_time_none.Height
        opt_past.Label = "Past"
        opt_past.State = 0
        opt_past.Tabstop = True
        opt_past.TabIndex = 3
        opt_past_wrappper = ControlWrapper(
            opt_past.Name, opt_past, opt_past.PositionX, opt_past.PositionY
        )
        self._m_tab_syntax.insertByName(opt_past.Name, opt_past)
        self._m_opt_time_past = opt_past

        opt_future: UnoControlRadioButtonModel = self._m_tab_syntax.createInstance(
            UnoControlRadioButtonModel.__ooo_full_ns__
        )
        opt_future.Name = "optFuture"
        opt_future.PositionX = opt_time_none.PositionX
        opt_future.PositionY = opt_past_wrappper.bottom
        opt_future.Width = opt_time_none.Width
        opt_future.Height = opt_time_none.Height
        opt_future.Label = "Future"
        opt_future.State = 0
        opt_future.Tabstop = True
        opt_future.TabIndex = 4
        opt_future_wrappper = ControlWrapper(
            opt_future.Name, opt_future, opt_future.PositionX, opt_future.PositionY
        )
        self._m_tab_syntax.insertByName(opt_future.Name, opt_future)
        self._m_opt_time_future = opt_future

        gb_prefix_suffix: UnoControlGroupBoxModel = self._m_tab_syntax.createInstance(
            UnoControlGroupBoxModel.__ooo_full_ns__
        )
        gb_prefix_suffix.Name = "gbPrefixSuffix"
        gb_prefix_suffix.Width = self._get_tab_pixels_width(self._pg1_gb_width)
        gb_prefix_suffix.Height = self._get_tab_pixels_height(90)
        gb_prefix_suffix.PositionX = gb_time.PositionX
        gb_prefix_suffix.PositionY = (
            gb_time_wraper.bottom + self._get_tab_pixels_height(self._vert_sep)
        )
        gb_prefix_suffix.Label = "Suffix"
        gb_prefix_suffix.TabIndex = 5
        self._m_tab_syntax.insertByName(gb_prefix_suffix.Name, gb_prefix_suffix)
        self._m_gb_prefix_suffix = gb_prefix_suffix

        opt_pre_suf_none: UnoControlRadioButtonModel = (
            self._m_tab_syntax.createInstance(
                UnoControlRadioButtonModel.__ooo_full_ns__
            )
        )
        opt_pre_suf_none.Name = "optPrefixSuffixNone"
        opt_pre_suf_none.PositionX = (
            gb_prefix_suffix.PositionX + self._get_tab_pixels_width(self._horz_sep)
        )
        opt_pre_suf_none.PositionY = (
            gb_prefix_suffix.PositionY + self._get_tab_pixels_height(self._gb_y_ofset)
        )
        opt_pre_suf_none.Width = gb_prefix_suffix.Width - (
            self._get_tab_pixels_width(self._horz_sep) * 2
        )
        opt_pre_suf_none.Height = self._get_tab_pixels_height(self._radio_btn_height)
        opt_pre_suf_none.Label = "None"
        opt_pre_suf_none.State = 0
        opt_pre_suf_none.Tabstop = True
        opt_pre_suf_none.TabIndex = 6
        opt_pre_suf_none_wrappper = ControlWrapper(
            opt_pre_suf_none.Name,
            opt_pre_suf_none,
            opt_pre_suf_none.PositionX,
            opt_pre_suf_none.PositionY,
        )
        self._m_tab_syntax.insertByName(opt_pre_suf_none.Name, opt_pre_suf_none)
        self._m_opt_pre_suf_none = opt_pre_suf_none

        opt_prefix: UnoControlRadioButtonModel = self._m_tab_syntax.createInstance(
            UnoControlRadioButtonModel.__ooo_full_ns__
        )
        opt_prefix.Name = "optPrefix"
        opt_prefix.PositionX = opt_pre_suf_none.PositionX
        opt_prefix.PositionY = opt_pre_suf_none_wrappper.bottom
        opt_prefix.Width = opt_pre_suf_none.Width
        opt_prefix.Height = opt_pre_suf_none.Height
        opt_prefix.Label = "NC"
        opt_prefix.State = 0
        opt_prefix.Tabstop = True
        opt_prefix.TabIndex = 7
        opt_prefix_wrappper = ControlWrapper(
            opt_prefix.Name, opt_prefix, opt_prefix.PositionX, opt_prefix.PositionY
        )
        self._m_tab_syntax.insertByName(opt_prefix.Name, opt_prefix)
        self._m_opt_prefix = opt_prefix

        opt_suffix: UnoControlRadioButtonModel = self._m_tab_syntax.createInstance(
            UnoControlRadioButtonModel.__ooo_full_ns__
        )
        opt_suffix.Name = "optSuffix"
        opt_suffix.PositionX = opt_pre_suf_none.PositionX
        opt_suffix.PositionY = opt_prefix_wrappper.bottom
        opt_suffix.Width = opt_pre_suf_none.Width
        opt_suffix.Height = opt_pre_suf_none.Height
        opt_suffix.Label = "Suffix"
        opt_suffix.State = 0
        opt_suffix.Tabstop = True
        opt_suffix.TabIndex = 8
        opt_suffix_wrappper = ControlWrapper(
            opt_suffix.Name, opt_suffix, opt_suffix.PositionX, opt_suffix.PositionY
        )
        self._m_tab_syntax.insertByName(opt_suffix.Name, opt_suffix)
        self._m_opt_suffix = opt_suffix

    def _create_tab_void_controls(self) -> None:
        lbl_void_key: UnoControlFixedTextModel = self._m_tab_syntax.createInstance(
            UnoControlFixedTextModel.__ooo_full_ns__
        )
        lbl_void_key.PositionX = 0
        lbl_void_key.PositionY = 2
        lbl_void_key.Width = self._get_tab_pixels_width(self._m_tabs_model.Width)
        lbl_void_key.Height = self._get_tab_pixels_height(self._btn_height)
        lbl_void_key.Label = "VOID-OPTION"
        lbl_void_key.NoLabel = True
        lbl_void_key.Align = 1
        lbl_syntax_key_wraper = ControlWrapper(
            "lblSyntaxKey", lbl_void_key, lbl_void_key.PositionX, lbl_void_key.PositionY
        )
        self._m_tab_void.insertByName(lbl_syntax_key_wraper.name, lbl_void_key)
        self._m_lbl_void = lbl_void_key
        # rgb_bg = ucolor.rgb.from_hex("272727")

        lb_void_key: UnoControlListBoxModel = self._m_tab_syntax.createInstance(
            UnoControlListBoxModel.__ooo_full_ns__
        )
        lb_void_key.Name = "lbVoidKey"
        lb_void_key.PositionY = lbl_syntax_key_wraper.bottom + 2
        lb_void_key.Dropdown = False
        lb_void_key.Border = 2
        lb_void_key.Tabstop = True
        lb_void_key.TabIndex = 0
        lb_void_key.Width = self._get_tab_pixels_width(170)
        lb_void_key.Height = self._get_tab_pixels_height(160)
        posx = int(self._get_tab_pixels_width(self._m_tabs_model.Width) / 2) - int(
            lb_void_key.Width / 2
        )
        lb_void_key.PositionX = posx
        # lb_void_key.BackgroundColor = rgb_bg.to_int()
        self._m_tab_void.insertByName(lb_void_key.Name, lb_void_key)
        self._m_lb_void_key = lb_void_key

    def _create(self):
        self._builder = DialogBuilder(self._title, self._width, self._height)
        lbl_title: UnoControlFixedTextModel = self._builder.create_control(
            UnoControlFixedTextModel
        )
        lbl_title.Name = "lblTitle"
        lbl_title.Label = "Syntax"
        lbl_title.NoLabel = True
        # lbl_title.TextColor = Color(int("9fa0a0", 16))
        lbl_title.Align = 1
        lbl_title.PositionX = self._border_padding_x
        lbl_title.PositionY = 2
        lbl_title.Width = self._width - (self._border_padding_x * 2)
        lbl_title.Height = self._lbl_title_height
        lbl_title_wrapper = self._builder.add_control(lbl_title.Name, lbl_title)
        self._m_lbl_title = lbl_title
        self._m_lbl_title_wrapper = lbl_title_wrapper

        btn_ok: UnoControlButtonModel = self._builder.create_control(
            UnoControlButtonModel
        )
        btn_ok.Name = self._name_btn_ok
        btn_ok.Label = "OK"
        btn_ok.TabIndex = 1
        btn_ok.DefaultButton = True
        btn_ok.Width = 100
        btn_ok.Height = self._btn_height
        btn_ok.PositionX = self._width - (btn_ok.Width + self._border_padding_x)
        btn_ok.PositionY = self._height - btn_ok.Height - 5
        # btn_ok.PushButtonType = PushButtonType.OK
        self._builder.add_control(btn_ok.Name, btn_ok)
        self._m_btn_ok = btn_ok

        btn_cancel: UnoControlButtonModel = self._builder.create_control(
            UnoControlButtonModel
        )
        btn_cancel.Name = self._name_btn_cancel
        btn_cancel.TabIndex = 2
        btn_cancel.DefaultButton = False
        btn_cancel.Width = btn_ok.Width
        btn_cancel.Height = self._btn_height
        btn_cancel.PositionX = self._border_padding_x
        btn_cancel.PositionY = btn_ok.PositionY
        btn_cancel.PushButtonType = PushButtonType.CANCEL
        self._builder.add_control(btn_cancel.Name, btn_cancel)
        self._m_btn_cancel = btn_cancel

        self._create_tabs()

        # init must be call before tabs_model,insertByIndex() is called
        self._builder.init()

        self._m_tabs_model.insertByIndex(0, self._m_tab_syntax)
        self._m_tabs_model.insertByIndex(1, self._m_tab_void)
        self.active_tab = TabEnum.SYNTAX

        window = self._builder.get_window()
        x: int = 0
        y: int = 0
        if window:
            ps = window.getPosSize()
            x = ps.Width / 2 - self._width / 2
            y = ps.Height / 2 - self._height / 2
        else:
            raise Exception("Unable to obtain size of window")
        self._builder.set_position_size(x, y)

    # endregion Create Dialog

    # region Callbacks

    def _ok(self) -> None:
        self._controller.handle_click_ok()

    def _cancel(self) -> None:
        self.exit()

    def _tab_changed(self, tab_page_activated_event: "TabPageActivatedEvent") -> None:
        try:
            self._controller.selected_tab = TabEnum(tab_page_activated_event.TabPageID)
        except Exception as e:
            print(e)
            raise

    def _dpv_state_change(self, evt: "PropertyChangeEvent") -> None:
        try:
            self._controller.controller_syntax.dpv = bool(evt.NewValue)
        except Exception as e:
            print(e)
            raise

    def _opt_time_state_change(self, evt: "PropertyChangeEvent") -> None:
        # state (evn.NewValue) will be 1 for true and 0 for false
        try:
            src: UnoControlRadioButtonModel = evt.Source
            if src.Name == self._m_opt_time_none.Name:
                self._controller.controller_syntax.time = TimeEnum.NONE
            elif src.Name == self._m_opt_time_past.Name:
                self._controller.controller_syntax.time = TimeEnum.PAST
            else:
                self._controller.controller_syntax.time = TimeEnum.FUTURE
        except Exception as e:
            print(e)
            raise

    def _opt_pre_suf_state_change(self, evt: "PropertyChangeEvent") -> None:
        try:
            # state (evn.NewValue) will be 1 for true and 0 for false
            src: UnoControlRadioButtonModel = evt.Source
            if src.Name == self._m_opt_prefix.Name:
                self._controller.controller_syntax.text = TextEnum.PREFIX
            elif src.Name == self._m_opt_suffix.Name:
                self._controller.controller_syntax.text = TextEnum.SUFFIX
            else:
                self._controller.controller_syntax.text = TextEnum.NONE
        except Exception as e:
            print(e)
            raise

    def _lb_syntax_change(self, evt: "ItemEvent"):

        # self._builder.dialog.getControl("lbSyntaxKey").SelectedItems[0]
        # the above line get the text of the selected item, the below line gets the index
        # self._builder.model.lbSyntaxKey.SelectedItems[0]
        # When using getControl() getSelectedItemPos() returns index and getSelectedItem() gets str
        try:
            lb = self.ctl_lb_syntax
            selected_item = lb.getSelectedItem()
            self._controller.controller_syntax.syntax = SyntaxEnum(selected_item)
        except Exception as e:
            print(e)
            raise

    def _lb_void_change(self, evt: "ItemEvent"):
        try:
            lb = self.ctl_lb_void
            selected_item = lb.getSelectedItem()
            self._controller.controller_void.void = VoidEnum(selected_item)
        except Exception as e:
            print(e)
            raise

    # endregion Callbacks

    # region UI Commands
    def exit(self):
        """
        Terminates this dialog box
        """
        self._builder.dialog.endExecute()

    # endregion UI Commands

    # region UI UPDATES

    def _set_dialog_properties(self):
        try:
            print("Setting Dialog Properties")
            if self._controller.selected_tab == TabEnum.SYNTAX:
                self._m_lbl_title.Label = self._controller.selected_text
                if self._controller.controller_syntax.is_syntax():
                    self.ctl_lb_syntax.selectItem(
                        self._controller.controller_syntax.syntax.name, True
                    )
                    if self._controller.controller_syntax.time == TimeEnum.PAST:
                        self._m_opt_time_past.State = True
                    elif self._controller.controller_syntax.time == TimeEnum.FUTURE:
                        self._m_opt_time_future.State = True
                    elif self._controller.controller_syntax.time == TimeEnum.NONE:
                        self._m_opt_time_none.State = True

                    if self._controller.controller_syntax.text == TextEnum.PREFIX:
                        self._m_opt_prefix.State = True
                    elif self._controller.controller_syntax.text == TextEnum.SUFFIX:
                        self._m_opt_suffix.State = True
                    elif self._controller.controller_syntax.text == TextEnum.NONE:
                        self._m_opt_pre_suf_none.State = True
                    if self._controller.controller_syntax.dpv == True:
                        self._m_chk_dpv.State = 1
                else:
                    self._m_opt_time_none.State = True
                    self._m_opt_pre_suf_none.State = True
            else:
                self._m_lbl_title.Label = self._controller.selected_text
                if self._controller.controller_void.void is not None:
                    self.ctl_lb_void.selectItem(
                        self._controller.controller_void.void.name, True
                    )
            print("Finished Setting Dialog Properties")
        except Exception as e:
            print(e)
            raise

    def refresh(self):
        if self._refreshing:
            return
        print("Refreshing")
        self._refreshing = True
        try:
            if self._controller.selected_tab == TabEnum.SYNTAX:
                if self._controller.controller_syntax.is_syntax():
                    self._m_chk_dpv.Enabled = (
                        self._controller.controller_syntax.syntax == SyntaxEnum.ORDER
                    )
                    self._m_gb_time.Enabled = (
                        self._controller.controller_syntax.syntax < SyntaxEnum.PLAN
                    )
                    self._m_gb_prefix_suffix.Enabled = (
                        self._controller.controller_syntax.syntax < SyntaxEnum.METHOD
                    )
                    self._m_btn_ok.Enabled = True
                else:
                    self._m_chk_dpv.Enabled = False
                    self._m_gb_time.Enabled = False
                    self._m_gb_prefix_suffix.Enabled = False
                    self._m_btn_ok.Enabled = False

                self._m_opt_time_none.Enabled = self._m_gb_time.Enabled
                self._m_opt_time_past.Enabled = self._m_gb_time.Enabled
                self._m_opt_time_future.Enabled = self._m_gb_time.Enabled

                self._m_opt_pre_suf_none.Enabled = self._m_gb_prefix_suffix.Enabled
                self._m_opt_prefix.Enabled = self._m_gb_prefix_suffix.Enabled
                self._m_opt_suffix.Enabled = self._m_gb_prefix_suffix.Enabled
            else:
                is_selection = not self._controller.controller_void.void is None
                self._m_btn_ok.Enabled = is_selection
            self._set_dialog_properties()
        except Exception as e:
            print(e)
            raise
        finally:
            self._refreshing = False
        print("Refresh finished")

    # endregion UI UPDATES

    # region Properties
    @property
    def dialog(self) -> "UnoControlDialog":
        """
        Gets this instance of dialog

        Returns:
            object: dialog object
        """
        return self._builder.dialog

    @property
    def model(self) -> object:
        """
        gets instance model

        Returns:
            object: model object
        """
        return self._builder.model

    @property
    def active_tab(self) -> TabEnum:
        """Gets/sets activte tab"""
        return TabEnum(self.ctl_tab.ActiveTabPageID)

    @active_tab.setter
    def active_tab(self, value: TabEnum) -> None:
        self.ctl_tab.ActiveTabPageID = value.value

    # endregion Properties

    # region Control Properties

    @property
    def ctl_btn_ok(self) -> "UnoControlButton":
        """Gets ctl_lbl_title value"""
        return self.dialog.getControl(self._m_btn_ok.Name)

    @property
    def ctl_btn_cancel(self) -> "UnoControlButton":
        """Gets ctl_lbl_title value"""
        return self.dialog.getControl(self._m_btn_cancel.Name)

    @property
    def ctl_tab(self) -> "UnoControlTabPageContainer":
        return self.dialog.getControl(self._m_tabs_model.Name)

    @property
    def ctl_tab_syntax(self) -> "UnoControlTabPage":
        return self.ctl_tab.getTabPage(0)

    @property
    def ctl_tab_void(self) -> "UnoControlTabPage":
        return self.ctl_tab.getTabPage(1)

    @property
    def ctl_chk_dpv(self) -> "UnoControlCheckBox":
        return self.ctl_tab_syntax.getControl(self._m_chk_dpv.Name)

    @property
    def ctl_opt_time_none(self) -> "UnoControlRadioButton":
        return self.ctl_tab_syntax.getControl(self._m_opt_time_none.Name)

    @property
    def ctl_opt_time_past(self) -> "UnoControlRadioButton":
        return self.ctl_tab_syntax.getControl(self._m_opt_time_past.Name)

    @property
    def ctl_opt_time_future(self) -> "UnoControlRadioButton":
        return self.ctl_tab_syntax.getControl(self._m_opt_time_future.Name)

    @property
    def ctl_opt_pre_suf_none(self) -> "UnoControlRadioButton":
        return self.ctl_tab_syntax.getControl(self._m_opt_pre_suf_none.Name)

    @property
    def ctl_opt_pre_suf(self) -> "UnoControlRadioButton":
        return self.ctl_tab_syntax.getControl(self._m_opt_pre_suf_none.Name)

    @property
    def ctl_opt_prefix(self) -> "UnoControlRadioButton":
        return self.ctl_tab_syntax.getControl(self._m_opt_prefix.Name)

    @property
    def ctl_opt_suffix(self) -> "UnoControlRadioButton":
        return self.ctl_tab_syntax.getControl(self._m_opt_suffix.Name)

    @property
    def ctl_lb_syntax(self) -> "UnoControlListBox":
        return self.ctl_tab_syntax.getControl(self._m_lb_syntax.Name)

    @property
    def ctl_lb_void(self) -> "UnoControlListBox":
        return self.ctl_tab_void.getControl(self._m_lb_void_key.Name)

    @property
    def ctl_gb_time(self) -> "UnoControlGroupBox":
        return self.ctl_tab_syntax.getControl(self._m_gb_time.Name)

    @property
    def ctl_gb_pre_suf(self) -> "UnoControlGroupBox":
        return self.ctl_tab_syntax.getControl(self._m_gb_prefix_suffix.Name)

    @property
    def ctl_lbl_title(self) -> "UnoControlFixedText":
        return self.dialog.getControl(self._m_lbl_title.Name)

    # endregion Control Properties
