"""
Example of a dialog that allows the user to rotate shapes in a Draw document.

For documentation see: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/macro/rotate_shape
"""

from __future__ import annotations

from typing import Any, cast, TYPE_CHECKING
from ooodev.draw import DrawDoc
from ooodev.utils.color import StandardColor
from ooodev.dialog.msgbox import MessageBoxResultsEnum, MessageBoxType
from ooodev.dialog import BorderKind
from ooodev.loader import Lo
from ooodev.events.args.event_args import EventArgs
from ooo.dyn.awt.font_descriptor import FontDescriptor
from ooodev.utils.info import Info
from ooo.dyn.awt.push_button_type import PushButtonType
from ooo.dyn.awt.pos_size import PosSize
from ooodev.loader.inst.doc_type import DocType
from ooodev.adapter.drawing.rotation_descriptor_properties_partial import (
    RotationDescriptorPropertiesPartial,
)
from ooodev.units import Angle100


if TYPE_CHECKING:
    from ooodev.dialog.dl_control.ctl_radio_button import CtlRadioButton


class RotateDialog:
    """
    Creates a dialog and adds controls to it.

    The dialog offers options to rotate shapes in a Draw document.
    """

    def __init__(self) -> None:
        self._doc = Lo.current_doc
        self._border_kind = BorderKind.BORDER_SIMPLE
        self._width = 400
        self._height = 210
        self._btn_width = 100
        self._btn_height = 30
        self._margin = 6
        self._box_height = 30
        self._title = "Rotate Shapes"
        self._msg = "Rotate the selected shapes"
        if self._border_kind != BorderKind.BORDER_3D:
            self._padding = 10
        else:
            self._padding = 14
        self._current_tab_index = 1
        self._angle = 0.0
        self._group1_opt: CtlRadioButton | None = None
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

    def _get_horizontal_center(self, width: int) -> int:
        return round(self._width / 2)

    def _init_dialog(self) -> None:
        """Create dialog and add controls."""
        self._init_handlers()
        self._dialog = self._doc.create_dialog(
            x=-1, y=-1, width=self._width, height=self._height, title=self._title
        )
        self._dialog.set_visible(False)
        self._dialog.create_peer()
        self._init_buttons()
        self._init_label()
        self._init_group_box()
        self._init_radio_controls()
        self._init_num_ctl()

    def _init_handlers(self) -> None:
        """
        Add event handlers for when changes occur.

        Methods can not be assigned directly to control callbacks.
        This is a python thing. However, methods can be assigned to class
        variable an in turn those can be assigned to callbacks.
        """
        self._fn_on_group1_changed = self.on_group1_changed
        self._fn_on_text_changed = self.on_text_changed

    def _init_buttons(self) -> None:
        """Add OK and Cancel and buttons to dialog control"""
        self._ctl_btn_cancel = self._dialog.insert_button(
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
        self._ctl_btn_ok = self._dialog.insert_button(
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

    def _init_label(self) -> None:
        """Add a fixed text label to the dialog control"""
        self._ctl_main_lbl = self._dialog.insert_label(
            label=self._msg,
            x=self._margin,
            y=self._padding,
            width=self._width - (self._padding * 2),
            height=self._box_height,
        )
        self._ctl_main_lbl.set_font_descriptor(self._fd)
        self._ctl_main_lbl.font_descriptor.weight = 150  # make bold

    def _init_group_box(self) -> None:
        """Create a group box to hold radio buttons."""
        sz_lbl = self._ctl_main_lbl.view.getPosSize()

        mid = self._get_horizontal_center(self._width)

        self._ctl_gb1 = self._dialog.insert_group_box(
            x=self._margin + mid,
            y=self._padding + sz_lbl.X + sz_lbl.Height,
            height=100,
            width=round(mid - (self._margin * 2)),
            label="Rotation Options",
        )
        self._ctl_gb1.set_font_descriptor(self._fd)
        self._ctl_gb1.font_descriptor.weight = 150  # make bold

    def _init_radio_controls(self) -> None:
        """
        Inserts radio buttons into dialog.
        """
        sz_gb1 = self._ctl_gb1.view.getPosSize()
        self._rb1 = self._dialog.insert_radio_button(
            label="Rotate Absolute",
            x=sz_gb1.X + self._margin,
            y=sz_gb1.Y + self._box_height,
            width=sz_gb1.Width - (self._margin * 2),
            height=20,
            name="rb_absolute",
        )
        self._rb1.set_font_descriptor(self._fd)
        self._group1_opt = self._rb1
        self._ctl_gb1.model.TabIndex = self._current_tab_index
        self._current_tab_index += 1

        self._rb1.model.State = 1
        self._rb1.tab_index = self._current_tab_index
        self._current_tab_index += 1

        self._rb1.add_event_item_state_changed(self._fn_on_group1_changed)
        rb1_sz = self._rb1.view.getPosSize()

        self._rb2 = self._dialog.insert_radio_button(
            label="Rotate Relative",
            x=rb1_sz.X,
            y=rb1_sz.Y + rb1_sz.Height,
            width=rb1_sz.Width,
            height=rb1_sz.Height,
            name="rb_relative",
        )
        self._rb2.set_font_descriptor(self._fd)
        self._rb2.add_event_item_state_changed(self._fn_on_group1_changed)
        self._rb2.tab_index = self._current_tab_index
        self._current_tab_index += 1

    def _init_num_ctl(self) -> None:
        """Create a Number Control for the angle."""
        sz_gb1 = self._ctl_gb1.view.getPosSize()
        self._num_ctl = self._dialog.insert_numeric_field(
            x=50,
            y=sz_gb1.Y + 12,
            width=120,
            height=20,
            value=0,
            min_value=-360.0,
            max_value=360.0,
            spin_button=True,
        )
        self._num_ctl.set_font_descriptor(self._fd)
        self._num_ctl.add_event_text_changed(self._fn_on_text_changed)
        self._num_ctl.tab_index = self._current_tab_index
        self._num_ctl.repeat = True
        self._current_tab_index += 1

    # region Event Handlers
    def on_group1_changed(
        self, src: Any, event: EventArgs, control_src: CtlRadioButton, *args, **kwargs
    ) -> None:
        """Fires when the radio button selection changes."""
        # itm_event = cast("ItemEvent", event.event_data)
        self._group1_opt = control_src

    def on_text_changed(
        self, src: Any, event: EventArgs, control_src: Any, *args, **kwargs
    ) -> None:
        """Fires when the text in the numeric control changes."""
        self._angle = float(control_src.value)

    # endregion Event Handlers

    def show(self) -> bool:
        # make sure the document is active.
        self._doc.activate()
        window = self._doc.get_frame().getContainerWindow()
        ps = window.getPosSize()
        x = round(ps.Width / 2 - self._width / 2)
        y = round(ps.Height / 2 - self._height / 2)
        self._dialog.set_pos_size(x, y, self._width, self._height, PosSize.POSSIZE)
        self._dialog.set_visible(True)
        dialog_result = self._dialog.execute()
        result = False
        if dialog_result == MessageBoxResultsEnum.OK.value:
            result = True

        # self._handle_results(result)
        # self._dialog.dispose()
        self._dialog.end_dialog(0)
        return result

    @property
    def relative(self) -> bool:
        return self._group1_opt.name == "rb_relative"

    @property
    def angle(self) -> Angle100:
        # angle is in 100th of a degree
        # multiply by 100 to get the value in 100th of a degree
        return Angle100(round(self._angle * 100))


def rotate_dialog(shapes: list) -> Any:
    """
    This method creates a instance of the RotateDialog class and displays it.

    When the dialog is shown, the user can select the rotation type and the angle to rotate the selected shapes.

    If the user clicks the OK button, the shapes are rotated based on the user's selection.

    Each shapes is checked to see if it support the RotationDescriptorPropertiesPartial class.
    OooDev automatically injects RotationDescriptorPropertiesPartial into the shape class if is supports rotation.

    For the ``rotate_angle`` property, the value is in 100th of a degree. So, to rotate a shape 90 degrees, the value would be 9000.
    OooDev provides the Angle100 class to help with this. Getting the ``rotate_angle`` property will return an Angle100 object.
    Setting the ``rotate_angle`` property can be done with an Angle100 object or an integer value in  ``1/100th degrees``.

    Args:
        shapes (list): This is the list of shapes selected found in the current document.

    Returns:
        Any: _description_
    """
    dialog = RotateDialog()
    if dialog.show():
        # doc = Lo.current_doc
        # doc.msgbox(f"Relative: {dialog.relative}. Angle: {dialog.angle}")
        print(shapes)
        print("Hello")
        for shape in shapes:
            if isinstance(shape, RotationDescriptorPropertiesPartial):
                if dialog.relative:
                    # In OooDev rotate_angle can be set with an integer or with another angle unit
                    shape.rotate_angle += dialog.angle
                else:
                    shape.rotate_angle = dialog.angle


def rotate_selected_shapes(*args) -> None:
    """
    Displays a message box for Hello World

    This is the actual macro that will be called from the macro run dialog.
    """
    doc = Lo.current_doc
    if doc.DOC_TYPE != DocType.DRAW:
        doc.msgbox(
            "This macro only works with Draw documents.",
            title="Wrong Document",
            boxtype=MessageBoxType.WARNINGBOX,
        )
        return
    draw_doc = cast(DrawDoc, doc)
    selected = draw_doc.get_selected_shapes()
    if not selected:
        doc.msgbox(
            "No shapes are selected.",
            title="No Shapes Selected",
            boxtype=MessageBoxType.WARNINGBOX,
        )
        return

    rotate_dialog(selected)


g_exportedScripts = (rotate_selected_shapes,)
