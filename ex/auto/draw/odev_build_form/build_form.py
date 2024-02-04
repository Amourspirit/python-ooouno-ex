# region imports
from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
from pathlib import Path
import datetime
import contextlib

import uno  # noqa: F401
from com.sun.star.sdbc import XResultSet

from ooodev.adapter.awt.top_window_events import TopWindowEvents
from ooodev.events.args.event_args import EventArgs
from ooodev.form import BorderKind, TriStateKind
from ooodev.form import ListSourceType
from ooodev.format.writer.direct.char.font import Font
from ooodev.loader import Lo
from ooodev.units import UnitMM
from ooodev.utils.color import StandardColor
from ooodev.utils.file_io import FileIO
from ooodev.draw import DrawDoc, ZoomKind
from ooodev.format.draw.direct.position_size.position_size import Protect

from ooodev.format.draw.direct.area import (
    Gradient as DrawAreaGradient,
    PresetGradientKind,
)

if TYPE_CHECKING:
    from com.sun.star.awt import ItemEvent
    from com.sun.star.beans import PropertyChangeEvent
    from ooodev.form.controls import (
        FormCtlButton,
        FormCtlGrid,
        FormCtlListBox,
        FormCtlTextField,
    )
    from ooodev.form.controls.form_ctl_base import DialogControlBase

# endregion imports


class BuildForm:
    # region Init
    def __init__(self, db_path: Path) -> None:
        super().__init__()

        self._db_fnm = db_path
        self.closed = False
        self._tab_index = 1
        self._init_callbacks()

        loader = Lo.load_office(Lo.ConnectSocket())
        try:
            self._doc = DrawDoc.create_doc(loader=loader, visible=True)

            # Delay to let the doc become visible before zooming.
            Lo.delay(500)
            self._doc.zoom(ZoomKind.ZOOM_100_PERCENT)

            self._top_win_ev = TopWindowEvents(add_window_listener=True)
            self._top_win_ev.add_event_window_closing(self._fn_on_window_closing)
            with Lo.ControllerLock():
                # use a controller lock to lock screen updating.
                # This will cut down and screen flashing and add controls faster.
                self.create_form()
            # dispatch a command that will turn off design mode.
            Lo.dispatch_cmd("SwitchControlDesignMode")

        except Exception:
            Lo.close_office()
            raise

    def _init_callbacks(self) -> None:
        # Event handlers are defined as methods on the class.
        # However class methods are not callable by the event system.
        # The solution is to assign the method to class fields and use them to add the event callbacks.
        self._fn_FormCtlButton = self.on_btn_action_preformed
        self._fn_on_btn_action_approved = self.on_btn_action_approved
        self._fn_on_down = self.on_down
        self._fn_on_grid_column_changed = self.on_grid_column_changed
        self._fn_on_grid_selection_changed = self.on_grid_selection_changed
        self._fn_on_list_item_changed = self.on_list_item_changed
        self._fn_on_mouse_entered = self.on_mouse_entered
        self._fn_on_mouse_exit = self.on_mouse_exit
        self._fn_on_property_change_listener = self.on_property_change_listener
        self._fn_on_up = self.on_up
        self._fn_on_veto_change_listener = self.on_veto_change_listener
        self._fn_on_window_closing = self.on_window_closing
        self._fn_text_changed = self.on_text_changed

    # endregion Init

    # region Form Creation

    def create_form(self) -> None:
        # Form has four sections: text, command_button, list_box, grid_control

        font_colored = Font(color=StandardColor.PURPLE_DARK3)
        draw_page = self._doc.slides[0]
        border_left = draw_page.border_left
        border_top = draw_page.border_top
        x1 = border_left + 2
        x2 = x1 + 42
        y = border_top + 11
        height = 6
        width1 = 40
        width2 = 40
        field_y_space = UnitMM(2)

        self._add_form_shape()

        main_form = draw_page.forms.add_form("MainForm")

        self._ctl_lbl_first_name = main_form.insert_control_label(
            label="FIRSTNAME",
            x=x1,
            y=y,
            width=width1,
            height=height,
            styles=[font_colored],
        )

        self._ctl_txt_first_name = main_form.insert_control_text_field(
            x=x2,
            y=y,
            width=width2,
            height=height,
            border=BorderKind.BORDER_SIMPLE,
        )
        self._ctl_lbl_first_name.bind_to_control(self._ctl_txt_first_name)
        self._ctl_txt_first_name.add_event_text_changed(self._fn_text_changed)
        self._set_tab_index(self._ctl_txt_first_name)

        y = (
            self._ctl_txt_first_name.position.y
            + self._ctl_txt_first_name.size.height
            + field_y_space
        )
        self._ctl_lbl_last_name = main_form.insert_control_label(
            label="LASTNAME",
            x=x1,
            y=y,
            width=width1,
            height=height,
            styles=[font_colored],
        )

        self._ctl_txt_last_name = main_form.insert_control_text_field(
            x=x2,
            y=y,
            width=width2,
            height=height,
            border=BorderKind.BORDER_SIMPLE,
        )
        self._ctl_lbl_last_name.bind_to_control(self._ctl_txt_last_name)
        self._ctl_txt_last_name.add_event_text_changed(self._fn_text_changed)
        self._set_tab_index(self._ctl_txt_last_name)

        y += 14

        self._ctl_lbl_age = main_form.insert_control_label(
            label="AGE",
            x=x1,
            y=y,
            width=width1,
            height=height,
            styles=[font_colored],
        )

        self._ctl_num_age = main_form.insert_control_numeric_field(
            x=x2,
            y=y,
            width=width2,
            height=8,
            accuracy=0,
            spin_button=True,
            border=BorderKind.BORDER_3D,
            min_value=1,
        )
        self._ctl_lbl_age.bind_to_control(self._ctl_num_age)
        self._ctl_num_age.add_event_down(self._fn_on_down)
        self._ctl_num_age.add_event_up(self._fn_on_up)
        self._set_tab_index(self._ctl_num_age)

        y = self._ctl_num_age.position.y + self._ctl_num_age.size.height + field_y_space
        self._ctl_lbl_birth_date = main_form.insert_control_label(
            label="BIRTHDATE",
            x=x1,
            y=y,
            width=width1,
            height=height,
            styles=[font_colored],
        )

        min_date = datetime.datetime.now() - datetime.timedelta(days=365 * 100)
        max_date = datetime.datetime.now()
        self._ctl_date_birth_date = main_form.insert_control_date_field(
            x=x2,
            y=y,
            width=width2,
            height=height,
            min_date=min_date,
            max_date=max_date,
            border=BorderKind.BORDER_3D,
        )
        self._ctl_lbl_birth_date.bind_to_control(self._ctl_date_birth_date)
        self._set_tab_index(self._ctl_date_birth_date)

        # buttons, all with listeners
        col1_x = border_left + 2
        x = col1_x
        spacing = 10
        y += 10
        width = 8

        names, labels = map(
            list,
            zip(
                *[
                    ["BtnFirst", "<<"],
                    ["BtnPrev", "<"],
                    ["BtnNext", ">"],
                    ["btnLast", ">>"],
                    ["BtnNew", ">*"],
                ]
            ),
        )

        for i in range(len(labels)):
            btn = main_form.insert_control_button(
                x=x + i * spacing,
                y=y,
                width=width,
                name=names[i],
            )
            btn.label = labels[i]
            btn.add_event_action_performed(self._fn_FormCtlButton)
            btn.add_event_approve_action(self._fn_on_btn_action_approved)
            self._set_tab_index(btn)

        self._ctl_btn_reload = main_form.insert_control_button(
            x=x + 4 * spacing + 16,
            y=y,
            width=13,
            name="BtnReload",
        )
        self._ctl_btn_reload.label = "reload"
        self._ctl_btn_reload.add_event_action_performed(self._fn_FormCtlButton)
        self._ctl_btn_reload.add_event_approve_action(self._fn_on_btn_action_approved)
        self._ctl_btn_reload.add_event_mouse_entered(self._fn_on_mouse_entered)
        self._ctl_btn_reload.add_event_mouse_exited(self._fn_on_mouse_exit)
        self._set_tab_index(self._ctl_btn_reload)

        # some fixed text; no listener
        y += 20

        lbl_sales_since = main_form.insert_control_label(
            x=x,
            y=y,
            width=60,
            height=self._ctl_lbl_first_name.size.height,
            label="Show only sales since",
            styles=[font_colored],
        )

        #  radio buttons inside a group box; use a property change listener
        col2_x = border_left + 90
        box_width = 70
        y = self._ctl_txt_first_name.position.y - 5
        name = "Options"

        opt_gb = main_form.insert_control_group_box(
            x=col2_x,
            y=y,
            width=box_width,
            height=25,
            label="Options",
            styles=[font_colored],
        )

        # these three radio buttons have the same name ("Option"), and
        # so only one can be on at a time
        indent = 3
        x = opt_gb.position.x + indent
        width = opt_gb.size.width - (2 * indent)

        labels = [
            "No automatic generation",
            "Before inserting a record",
            "When moving to a new record",
        ]

        for i in range(len(labels)):
            radio_btn = main_form.insert_control_radio_button(
                name=name,
                # label=labels[i],
                x=x,
                y=y + (i + 1) * height,
                width=width,
                height=height,
                styles=[font_colored],
            )
            radio_btn.label = labels[i]
            radio_btn.add_event_property_change(
                "State", self._fn_on_property_change_listener
            )
            radio_btn.tag = f"RadioBtn{i+1}"
            self._set_tab_index(radio_btn)

        y += 30
        x = opt_gb.position.x
        width = opt_gb.size.width
        misc_gb = main_form.insert_control_group_box(
            x=x,
            y=y,
            width=box_width,
            height=25,
            label="Miscellaneous",
            styles=[font_colored],
        )

        y = misc_gb.position.y + 5
        width = width = misc_gb.size.width - (2 * indent)
        x = misc_gb.position.x + indent
        height = 6

        # check boxes inside another group box
        # use the same property change listener

        self._ctl_chk_default_date = main_form.insert_control_check_box(
            x=x,
            y=y,
            name="ChkDefaultDate",
            label='Default sales date to "today"',
            width=width,
            height=height,
            tri_state=False,
            state=TriStateKind.NOT_CHECKED,
            styles=[font_colored],
        )
        self._ctl_chk_default_date.help_text = (
            "When checked, newly entered sales records are pre-filled"
        )
        self._ctl_chk_default_date.add_event_property_change(
            "State", self._fn_on_property_change_listener
        )
        self._set_tab_index(self._ctl_chk_default_date)

        y += height
        self._ctl_chk_protect = main_form.insert_control_check_box(
            x=x,
            y=y,
            name="ChkProtect",
            label="Protect key fields from editing",
            width=width,
            height=height,
            tri_state=False,
            state=TriStateKind.NOT_CHECKED,
            styles=[font_colored],
        )
        self._ctl_chk_protect.help_text = "When checked, you cannot modify the values"
        self._ctl_chk_protect.add_event_property_change(
            "State", self._fn_on_property_change_listener
        )
        self._set_tab_index(self._ctl_chk_protect)

        y += height
        self._ctl_chk_empty = main_form.insert_control_check_box(
            x=x,
            y=y,
            name="ChkEmpty",
            label="Check for empty sales names",
            width=width,
            height=height,
            tri_state=False,
            state=TriStateKind.NOT_CHECKED,
            styles=[font_colored],
        )
        self._ctl_chk_empty.help_text = "When checked, you cannot enter empty values"
        self._ctl_chk_empty.add_event_property_change(
            "State", self._fn_on_property_change_listener
        )
        self._set_tab_index(self._ctl_chk_empty)

        # a list using simple text
        fruits = ("apple", "orange", "pear", "grape")
        width = 40
        height = 6
        y = lbl_sales_since.position.y + 10
        x = self._ctl_lbl_first_name.position.x
        self._ctl_lst_fruits = main_form.insert_control_list_box(
            x=x,
            y=y,
            name="LstFruits",
            entries=fruits,
            width=width,
            height=height,
            border=BorderKind.BORDER_3D,
        )
        self._ctl_lst_fruits.add_event_item_state_changed(self._fn_on_list_item_changed)
        self._set_tab_index(self._ctl_lst_fruits)

        # set Form's data source to be the DB_FNM database
        main_form.bind_form_to_table(
            src_name=FileIO.fnm_to_url(self._db_fnm), tbl_name="Course"
        )

        # a list filled using an SQL query on the form's data source
        x = border_left + 60

        self._ctl_db_lst_course_names = main_form.insert_db_control_list_box(
            x=x,
            y=y,
            name="LstCourseNames",
            width=width,
            height=height,
            multi_select=False,
            border=BorderKind.BORDER_3D,
        )
        self._ctl_db_lst_course_names.list_source_type = ListSourceType.SQL
        self._ctl_db_lst_course_names.list_source = ('SELECT "title" FROM "Course"',)
        self._ctl_db_lst_course_names.bound_column = 0
        self._ctl_db_lst_course_names.add_event_item_state_changed(
            self._fn_on_list_item_changed
        )
        self._set_tab_index(self._ctl_db_lst_course_names)

        # another list filled using a different SQL query on the form's data source
        # ------------------------ set up database grid/table ----------------
        # Align to right side of Misc group box
        x = (misc_gb.position.x + misc_gb.size.width) - width
        self._ctl_db_lst_stud_names = main_form.insert_db_control_list_box(
            x=x,
            y=y,
            name="LstStudNames",
            width=width,
            height=height,
            multi_select=False,
            border=BorderKind.BORDER_3D,
        )
        self._ctl_db_lst_stud_names.list_source_type = ListSourceType.SQL
        self._ctl_db_lst_stud_names.list_source = ('SELECT "lastName" FROM "Student"',)
        self._ctl_db_lst_stud_names.bound_column = 0
        self._ctl_db_lst_stud_names.add_event_item_state_changed(
            self._fn_on_list_item_changed
        )
        self._set_tab_index(self._ctl_db_lst_stud_names)

        # create a new form, GridForm,
        grid_form = draw_page.forms.add_form("GridForm")

        # which uses an SQL query as its data source
        grid_form.bind_form_to_sql(
            src_name=FileIO.fnm_to_url(self._db_fnm),
            cmd='SELECT "firstName", "lastName" FROM "Student"',
        )

        x = self._ctl_lbl_first_name.position.x
        # create the grid/table component and set its columns
        self._ctl_grid_sales_tbl = grid_form.insert_control_grid(
            name="GridSalesTable",
            x=x,
            y=self._ctl_lst_fruits.position.y + 12,
            width=100,
            height=40,
        )
        self._ctl_grid_sales_tbl.create_grid_column(
            data_field="firstName", col_kind="TextField", width=25
        )
        self._ctl_grid_sales_tbl.create_grid_column(
            data_field="lastName", col_kind="TextField", width=25
        )
        self._set_tab_index(self._ctl_grid_sales_tbl)
        self._ctl_grid_sales_tbl.add_event_selection_changed(
            self._fn_on_grid_selection_changed
        )
        self._ctl_grid_sales_tbl.add_event_column_changed(
            self._fn_on_grid_column_changed
        )

    # endregion Form Creation

    # region Add Forms Shape
    def _add_form_shape(self) -> None:
        dp = self._doc.slides[0]
        # get the page margins from the document styles
        # pos = Position(pos_x=0, pos_y=0)

        rect = dp.draw_rectangle(
            x=dp.border_left,
            y=dp.border_top,
            width=dp.width - (dp.border_left + dp.border_right),
            height=160,
        )
        # set the shape  to have a gradient fill
        gradient = DrawAreaGradient.from_preset(preset=PresetGradientKind.TEAL_BLUE)
        # set the shape to be positioned relative to the page
        protect = Protect(size=True, position=True)  # protect size and position
        rect.apply_styles(gradient, protect)

    # endregion Add Forms Shape

    # region Other Methods
    def _set_tab_index(self, ctl: DialogControlBase) -> None:
        ctl.tab_index = self._tab_index
        self._tab_index += 1

    # endregion Other Methods

    # region Event Listeners
    def on_btn_action_approved(
        self, src: Any, event: EventArgs, control_src: FormCtlButton, *args, **kwargs
    ) -> None:
        print(
            f"Action Approved: '{control_src.model.Label}', Control Name: {control_src.name}"
        )

    def on_btn_action_preformed(
        self, src: Any, event: EventArgs, control_src: FormCtlButton, *args, **kwargs
    ) -> None:
        print(
            f"Action Performed: '{control_src.model.Label}', Control Name: {control_src.name}"
        )

    def on_down(
        self, src: Any, event: EventArgs, control_src: Any, *args, **kwargs
    ) -> None:
        print("Down:", control_src.name)
        print("Value:", control_src.value)

    def on_grid_column_changed(
        self, src: Any, event: EventArgs, control_src: FormCtlGrid, *args, **kwargs
    ) -> None:
        # print(control_src)
        print("Grid Column change", control_src.name)

    def on_grid_selection_changed(
        self, src: Any, event: EventArgs, control_src: FormCtlGrid, *args, **kwargs
    ) -> None:
        # used by the grid/table control

        # in the form module
        # not the one in com.sun.star.awt.grid
        # no way to get the current row with XGridControl; see form.XGrid
        print(
            f"Grid {control_src.name} column: {control_src.get_current_column_position()}"
        )

        # must access the result set inside the form for row info
        form_name = control_src.get_form_name()
        if not form_name:
            return
        gform = self._doc.slides[0].forms[form_name]
        rs = gform.qi(XResultSet)
        if not rs:
            return
        try:
            print(f"    row: {rs.getRow()}")
        except Exception as e:
            print(e)

    def on_list_item_changed(
        self, src: Any, event: EventArgs, control_src: FormCtlListBox, *args, **kwargs
    ) -> None:
        itm_event = cast("ItemEvent", event.event_data)
        # print("Selected:", itm_event.Selected)
        # print("Highlighted:", itm_event.Highlighted)
        selected_item = control_src.get_item(itm_event.Selected)
        print(f'List: "{control_src.name}", selection: {selected_item}')

    def on_mouse_entered(
        self, src: Any, event: EventArgs, control_src: Any, *args, **kwargs
    ) -> None:
        # print(control_src)
        print("Mouse Entered:", control_src.name)

    def on_mouse_exit(
        self, src: Any, event: EventArgs, control_src: Any, *args, **kwargs
    ) -> None:
        # print(control_src)
        print("Mouse Exited:", control_src.name)

    def on_property_change_listener(
        self,
        src: Any,
        event: EventArgs,
        control_src: DialogControlBase,
        *args,
        **kwargs,
    ) -> None:
        ev_data = cast("PropertyChangeEvent", event.event_data)
        print("Property Change:", control_src.name)
        print("Old Value:", ev_data.OldValue)
        print("New Value:", ev_data.NewValue)
        with contextlib.suppress(AttributeError):
            tag = control_src.tag
            if tag:
                print("Tag:", tag)

    def on_text_changed(
        self, src: Any, event: EventArgs, control_src: FormCtlTextField, *args, **kwargs
    ) -> None:
        # ev_arg = cast("TextEvent", event.event_data)
        print(f"listen_to_text_field2: {control_src.name}: {control_src.text}")

    def on_up(
        self, src: Any, event: EventArgs, control_src: Any, *args, **kwargs
    ) -> None:
        print("Up:", control_src.name)
        print("Value:", control_src.value)

    def on_veto_change_listener(
        self, src: Any, event: EventArgs, control_src: Any, *args, **kwargs
    ) -> None:
        ev_data = cast("PropertyChangeEvent", event.event_data)
        print("Vetoable Change:", control_src.name)
        print("Old Value:", ev_data.OldValue)
        print("New Value:", ev_data.NewValue)

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

    # endregion Event Listeners
