# coding: utf-8
from ooo.dyn.awt.uno_control_button_model import UnoControlButtonModel
from ooo.dyn.awt.uno_control_fixed_text_model import UnoControlFixedTextModel
from ooo.dyn.awt.uno_control_edit_model import UnoControlEditModel
from ooo.dyn.awt.push_button_type import PushButtonType
from ooo.dyn.awt.selection import Selection
from typing import TYPE_CHECKING
from ...ui.dialog.builder.dialog_builder import DialogBuilder


if TYPE_CHECKING:
    from ooo.csslo.awt import UnoControlButton, UnoControlDialog, UnoControlEdit


class InputBox:
    """InputBox Class for generating and displaying input box via UNO."""

    def __init__(
        self, message: str, title: str = "", default: str = "", x: int = -1, y: int = -1
    ):
        """
        Inits Input Box

        Args:
            message (str): message message to show on the dialog
            title (str, optional): title window title. Defaults to "".
            default (str, optional): default value in dialog. Defaults to "".
            x (int, optional): dialog positio in twips. Defaults to -1.
            y (int, optional): dialog position in twips. Defaults to -1.
        """
        self._message: str = message
        self._title: str = title
        self._default: str = default
        self._x: int = x
        self._y: int = y

        self._width = 600
        self._hori_margin = 8
        self._vert_margin = self._hori_margin
        self._button_width = 100
        self._button_height = 26
        self._hori_sep = 8
        self._vert_sep = self._hori_sep
        self._label_height = self._button_height * 2 + 5
        self._edit_height = 24
        self._height = (
            self._vert_margin * 2
            + self._label_height
            + self._vert_sep
            + self._edit_height
        )
        self._label_width = (
            self._width - self._button_width - self._hori_sep - self._hori_margin * 2
        )

    def show(self) -> str:
        """
        Show this Input Box

        Returns:
            str: input result if OK button pushed; Otherwise, zero length string
        """
        self._create()
        edit = self.ctl_edit
        edit.setSelection(Selection(min=0, max=len(str(self._default))))
        edit.setFocus()
        ret = edit.getModel().Text if self.dialog.execute() else ""
        self.dialog.dispose()
        return ret

    def _create(self):
        self._builder = DialogBuilder(self._title, self._width, self._height)
        lbl_title = self._builder.create_control(UnoControlFixedTextModel)
        lbl_title.Name = "label"
        lbl_title.Label = str(self._message)
        lbl_title.NoLabel = True
        lbl_title.Align = 0
        lbl_title.PositionX = self._hori_margin
        lbl_title.PositionY = self._vert_margin
        lbl_title.Width = self._label_width
        lbl_title.Height = self._label_height
        self._m_lbl_title = lbl_title

        btn_ok = self._builder.create_control(UnoControlButtonModel)
        btn_ok.Name = "btn_ok"
        btn_ok.TabIndex = 2
        btn_ok.DefaultButton = False
        btn_ok.Width = self._button_width
        btn_ok.Height = self._button_height
        btn_ok.PositionX = self._hori_margin + self._label_width + self._hori_sep
        btn_ok.PositionY = self._vert_margin
        btn_ok.PushButtonType = PushButtonType.OK
        self._m_btn_ok = btn_ok

        btn_cancel: UnoControlButtonModel = self._builder.create_control(
            UnoControlButtonModel
        )
        btn_cancel.Name = "btn_cancel"
        btn_cancel.TabIndex = 1
        btn_cancel.DefaultButton = True
        btn_cancel.Width = btn_ok.Width
        btn_cancel.Height = self._btn_height
        btn_cancel.PositionX = self._border_padding_x
        btn_cancel.PositionY = btn_ok.PositionY
        btn_cancel.PushButtonType = PushButtonType.CANCEL
        self._m_btn_cancel = btn_cancel

        ctl_edit = self._builder.create_control(UnoControlEditModel)
        ctl_edit.Name = "edit"
        ctl_edit.PositionX = self._hori_margin
        ctl_edit.PositionY = self._label_height + self._vert_margin + self._vert_sep
        ctl_edit.Width = self._width - self._hori_margin * 2
        ctl_edit.Height = self._edit_height
        ctl_edit.Text = self._default
        self._m_ctl_edit = btn_cancel

        self._builder.add_control(self._m_lbl_title.Name, self._m_lbl_title)
        self._builder.add_control(self._m_btn_ok.Name, self._m_btn_ok)
        self._builder.add_control(self._m_btn_cancel.Name, self._m_btn_cancel)
        self._builder.add_control(self._m_ctl_edit.Name, self._m_ctl_edit)

        self._builder.init()

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

    # region Control Properties

    @property
    def ctl_btn_ok(self, Action) -> "UnoControlButton":
        """Gets ctl_lbl_title value"""
        return self.dialog.getControl(self._m_btn_ok.Name)

    @property
    def ctl_btn_cancel(self) -> "UnoControlButton":
        """Gets ctl_lbl_title value"""
        return self.dialog.getControl(self._m_btn_cancel.Name)

    @property
    def ctl_edit(self) -> "UnoControlEdit":
        return self.dialog.getControl(self._m_ctl_edit.Name)

    # endregion Control Properties

    # endregion Properties
