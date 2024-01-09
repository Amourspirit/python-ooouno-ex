from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno
from ooo.dyn.script.script_event_descriptor import ScriptEventDescriptor

from ooodev.utils.lo import Lo
from ooodev.utils.props import Props
from ooodev.calc import Calc, CalcDoc, ZoomKind
from ooodev.adapter.awt.top_window_events import TopWindowEvents
from ooodev.events.args.event_args import EventArgs
from ooodev.format.draw.direct.position_size.position_size import Protect
from ooodev.format.writer.direct.char.font import Font
from ooodev.utils.color import StandardColor
from ooodev.units import UnitMM
from ooodev.form import BorderKind
from ooodev.dialog.input import Input
from ooodev.format.draw.direct.area import (
    Gradient as DrawAreaGradient,
    PresetGradientKind,
)

if TYPE_CHECKING:
    from com.sun.star.awt import ItemEvent
    from ooodev.draw.shapes.rectangle_shape import RectangleShape
    from ooodev.calc import CalcSheet
    from ooodev.form.controls import (
        FormCtlButton,
        FormCtlGrid,
        FormCtlListBox,
        FormCtlTextField,
    )
    from com.sun.star.document import DocumentEvent
    from com.sun.star.sheet import ActivationEvent


class BuildForm:
    def __init__(self) -> None:
        super().__init__()

        self.closed = False
        self._tab_index = 1
        self._init_callbacks()

        loader = Lo.load_office(Lo.ConnectSocket())
        try:
            self._doc = CalcDoc(Calc.create_doc(loader))

            self._doc.set_visible()
            # Delay to let the doc become visible before zooming.
            Lo.delay(500)
            self._doc.zoom(ZoomKind.ZOOM_100_PERCENT)

            self._init_callbacks()
            self._top_win_ev = TopWindowEvents(add_window_listener=True)
            self._top_win_ev.add_event_window_closing(self._fn_on_window_closing)
            self._add_lookup_sheet()
            with Lo.ControllerLock():
                # use a controller lock to lock screen updating.
                # This will cut down and screen flashing and add controls faster.
                self._create_form()
            # dispatch a command that will turn off design mode.
            Lo.dispatch_cmd("SwitchControlDesignMode")
            # self._doc.sheets[0].dispatch_recalculate()
            self._doc.add_event_document_event_occurred(self._fn_on_document_event)
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
        self._fn_on_lookup_modified = self.on_lookup_modified
        self._fn_on_window_closing = self.on_window_closing
        self._fn_on_list_item_changed = self.on_list_item_changed
        self._fn_on_document_event = self.on_document_event
        self._fn_on_active_sheet_changed = self.on_active_sheet_changed
        self._fn_text_changed = self.on_text_changed

    def _add_event_desc(self) -> None:
        # https://ask.libreoffice.org/t/solved-is-it-possible-to-insert-buttons-into-a-spreadsheet-from-a-macro/94771/7

        desc = ScriptEventDescriptor()
        desc.AddListenerParam = ""
        desc.EventMethod = "actionPerformed"
        desc.ListenerType = "XActionListener"  # "com.sun.star.awt.XActionListener"
        desc.ScriptCode = "vnd.sun.star.script:build_form.BuildForm.ButtonClick?language=Python&location=document"

    def _create_form(self) -> None:
        draw_page = self._doc.sheets[0].draw_page
        rect = self._add_form_shape()
        font_colored = Font(color=StandardColor.PURPLE_DARK3)

        pos = rect.position
        border_top = pos.y + 2
        border_left = pos.x + 2
        x = border_left
        height = UnitMM(6)
        width_lbl = UnitMM(24)
        width_box = UnitMM(40)
        field_y_space = UnitMM(2)
        field_x_space = UnitMM(2)
        main_form = draw_page.forms.add_form("MainForm")
        y = border_top

        self._ctl_lbl_first_name = main_form.insert_control_label(
            label="FIRSTNAME",
            x=x,
            y=y,
            width=width_lbl,
            height=height,
            styles=[font_colored],
        )

        x = (
            self._ctl_lbl_first_name.position.x
            + self._ctl_lbl_first_name.size.width
            + field_x_space
        )

        self._ctl_txt_first_name = main_form.insert_control_text_field(
            x=x,
            y=y,
            width=width_box,
            height=height,
            border=BorderKind.BORDER_SIMPLE,
        )
        self._ctl_lbl_first_name.bind_to_control(self._ctl_lbl_first_name)
        self._ctl_txt_first_name.add_event_text_changed(self._fn_text_changed)

        x = self._ctl_lbl_first_name.position.x
        y = (
            self._ctl_txt_first_name.position.y
            + self._ctl_txt_first_name.size.height
            + field_y_space
        )
        self._ctl_lbl_last_name = main_form.insert_control_label(
            label="LASTNAME",
            x=x,
            y=y,
            width=width_lbl,
            height=height,
            styles=[font_colored],
        )

        x = self._ctl_txt_first_name.position.x
        self._ctl_txt_last_name = main_form.insert_control_text_field(
            x=x,
            y=y,
            width=width_box,
            height=height,
            border=BorderKind.BORDER_SIMPLE,
        )
        self._ctl_lbl_last_name.bind_to_control(self._ctl_txt_last_name)
        self._ctl_txt_last_name.add_event_text_changed(self._fn_text_changed)

        y += field_x_space + height
        x = self._ctl_lbl_first_name.position.x
        self._ctl_lbl_age = main_form.insert_control_label(
            label="AGE",
            x=x,
            y=y,
            width=width_lbl,
            height=height,
            styles=[font_colored],
        )

        x = self._ctl_txt_first_name.position.x
        self._ctl_num_age = main_form.insert_control_numeric_field(
            x=x,
            y=y,
            width=width_box,
            height=8,
            accuracy=0,
            spin_button=True,
            border=BorderKind.BORDER_3D,
            min_value=1,
        )
        self._ctl_lbl_age.bind_to_control(self._ctl_num_age)

        y += field_x_space + self._ctl_num_age.size.height
        x = self._ctl_lbl_first_name.position.x
        self._ctl_lbl_choice = main_form.insert_control_label(
            label="CHOICE",
            x=x,
            y=y,
            width=width_lbl,
            height=height,
            styles=[font_colored],
        )

        x = self._ctl_txt_first_name.position.x
        self._ctl_lst_choices = main_form.insert_control_list_box(
            x=x,
            y=y,
            name="Choices",
            # entries=self._get_choices(),
            width=width_box,
            height=height,
            border=BorderKind.BORDER_3D,
        )
        # choices = self._get_choices()
        # choices.append("New Choice")
        # self._ctl_lst_choices.set_list_data(choices)
        self._ctl_lst_choices.add_event_item_state_changed(
            self._fn_on_list_item_changed
        )
        self._reload_choices()

    def _reload_choices(self) -> None:
        try:
            print("Reloading choices")
            self._ctl_lst_choices.set_list_data(self._get_choices())
            self._ctl_lst_choices.selected_items = (0,)
            print("Adding Listener")
            print("Reloaded choices")
        except Exception as e:
            print(f"Error reloading choices: {e}")

    def _add_new_choice(self) -> None:
        # get the new choice from the user
        new_choice = Input.get_input(title="New Choice", msg="Enter a new choice")
        # if the user canceled, then return
        if not new_choice:
            return
        sheet = self._doc.sheets["Lookup"]
        # get the used range
        rng = sheet.find_used_range_obj()
        # get the first column
        col_rng = rng.get_col(rng.col_start)
        # get the next new row
        cell_obj = col_rng.cell_end + 1
        # set the new choice in the sheet
        sheet[cell_obj].set_val(new_choice)

    def _get_choices(self) -> list[str]:
        sheet = self._doc.sheets["Lookup"]
        rng = sheet.find_used_range_obj()
        print(f"Lookup range: {rng}")
        col_rng = rng.get_col(rng.col_start)
        data = sheet.get_array(range_obj=col_rng)
        values = [i[0] for i in data if i[0]]
        val_set = set(values)  # remove duplicates
        result = list(val_set)
        result.sort()
        result.insert(0, "<< Add New >>")
        result.insert(0, "<< Select >>")
        print(f"Choices: {result}")
        return result

    def _add_form_shape(self) -> RectangleShape[CalcSheet]:
        dp = self._doc.sheets[0].draw_page
        # get the page margins from the document styles
        # pos = Position(pos_x=0, pos_y=0)

        rect = dp.draw_rectangle(
            x=100,
            y=10,
            width=100,
            height=130,
        )
        # set the shape  to have a gradient fill
        gradient = DrawAreaGradient.from_preset(preset=PresetGradientKind.TEAL_BLUE)
        # set the shape to be positioned relative to the page
        protect = Protect(size=True, position=True)  # protect size and position
        rect.apply_styles(gradient, protect)
        return rect

    def _add_lookup_sheet(self) -> None:
        # add a sheet to the document
        sheet = self._doc.sheets.insert_sheet(name="Lookup")

        # set the values for the lookup that can be used in the Choices list box.
        choice = [[f"Choice {i}"] for i in range(1, 9)]
        sheet.set_array(values=choice, name="A1")
        # add a listener to this sheet.
        # When the sheet is modified that will trigger the event.
        # When the event is triggered it will call the on_lookup_modified method.
        # the on_lookup_modified method will reload the choices in the list box.
        sheet.add_event_modified(self._fn_on_lookup_modified)

    def _show_recent_functions(self) -> None:
        recent_ids = Calc.get_recent_functions()
        if not recent_ids:
            return

        print(f"Recently used functions {len(recent_ids)}")
        for i in recent_ids:
            p = Calc.find_function(idx=i)
            print(f'  {Props.get_value(name="Name", props=p)}')

        print()

    # region Events

    # region Lookup Sheet Events
    def on_lookup_modified(
        self, source: Any, event_args: EventArgs, *args, **kwargs
    ) -> None:
        print("Lookup sheet modified")
        print("  Source:", source)
        print("  self:", self)
        if hasattr(self, "_ctl_lst_choices"):
            self._reload_choices()

    def on_list_item_changed(
        self, src: Any, event: EventArgs, control_src: FormCtlListBox, *args, **kwargs
    ) -> None:
        itm_event = cast("ItemEvent", event.event_data)
        # print("Selected:", itm_event.Selected)
        # print("Highlighted:", itm_event.Highlighted)
        selected_item = control_src.get_item(itm_event.Selected)
        print(f'List: "{control_src.name}", selection: {selected_item}')
        print("Selected Items:", self._ctl_lst_choices.model.SelectedItems)
        selected_index = self._ctl_lst_choices.model.SelectedItems[0]
        print(f"Selected index: {selected_index}")
        if selected_index == 1:
            self._add_new_choice()

    # endregion Lookup Sheet Events

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
    def on_document_event(
        self, source: Any, event_args: EventArgs, *args, **kwargs
    ) -> None:
        print("Document Event")
        try:
            event = cast("DocumentEvent", event_args.event_data)
            print("  Event:", event)
            print("    EventName:", event.EventName)
            print("    Source:", event.Source)
            print("    Document:", event.Supplement)

        except Exception as e:
            print(f"  {e}")

    def on_active_sheet_changed(
        self, source: Any, event_args: EventArgs, *args, **kwargs
    ) -> None:
        print("Active Sheet Changed")
        try:
            event = cast("ActivationEvent", event_args.event_data)
            # print("  Event:", event)
            # print("    ActiveSheet:", event.ActiveSheet)
            # print("    Source:", event.Source)
            sheet = self._doc.sheets.get_sheet(event.ActiveSheet)
            print("    Active Sheet:", sheet.name)
            if sheet.name == "Sheet1":
                self._ctl_lst_choices.add_event_item_state_changed(
                    self._fn_on_list_item_changed
                )
        except Exception as e:
            print(f"  {e}")

    def on_text_changed(
        self, src: Any, event: EventArgs, control_src: FormCtlTextField, *args, **kwargs
    ) -> None:
        # ev_arg = cast("TextEvent", event.event_data)
        print(f"listen_to_text_field2: {control_src.name}: {control_src.text}")

    # endregion Events
