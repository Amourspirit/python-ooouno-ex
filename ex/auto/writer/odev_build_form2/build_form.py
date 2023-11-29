# region imports
from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
from pathlib import Path
import datetime
import contextlib

import uno
import unohelper
from com.sun.star.awt import ActionEvent
from com.sun.star.awt import FocusEvent
from com.sun.star.awt import ItemEvent
from com.sun.star.awt import MouseEvent
from com.sun.star.awt import TextEvent
from com.sun.star.awt import XActionListener
from com.sun.star.awt import XButton
from com.sun.star.awt import XControlModel
from com.sun.star.awt import XFocusListener
from com.sun.star.awt import XItemListener
from com.sun.star.awt import XListBox
from com.sun.star.awt import XMouseListener
from com.sun.star.awt import XTextComponent
from com.sun.star.awt import XTextListener
from com.sun.star.awt import XWindow
from com.sun.star.beans import PropertyChangeEvent
from com.sun.star.beans import XPropertyChangeListener
from com.sun.star.beans import XPropertySet
from com.sun.star.document import XEventListener
from com.sun.star.form import XForm
from com.sun.star.form import XGridControl
from com.sun.star.form import XGridControlListener
from com.sun.star.lang import EventObject
from com.sun.star.sdbc import XResultSet
from com.sun.star.text import XTextDocument
from com.sun.star.view import XSelectionChangeListener
from com.sun.star.view import XSelectionSupplier

from ooodev.dialog.msgbox import (
    MsgBox,
    MessageBoxType,
    MessageBoxButtonsEnum,
    MessageBoxResultsEnum,
)
from ooodev.format.writer.direct.char.font import Font
from ooodev.office.write import Write
from ooodev.theme import ThemeGeneral
from ooodev.utils import color as color_util
from ooodev.utils.file_io import FileIO
from ooodev.utils.forms import Forms, FormComponentKind, FormComponentType
from ooodev.utils.gui import GUI
from ooodev.utils.info import Info
from ooodev.utils.lo import Lo
from ooodev.utils.props import Props
from ooodev.form import BorderKind, TriStateKind
from ooodev.events.args.event_args import EventArgs
from ooodev.adapter.awt.top_window_events import TopWindowEvents
from ooodev.form import ListSourceType

if TYPE_CHECKING:
    from com.sun.star.beans import PropertyChangeEvent
    from com.sun.star.awt import ItemEvent
    from ooodev.form.controls import (
        FormCtlFixedText,
        FormCtlGroupBox,
        FormCtlGrid,
        FormCtlTextField,
        FormCtlButton,
        FormCtlListBox,
    )
    from ooodev.form.controls.form_ctl_base import DialogControlBase
    from com.sun.star.awt import TextEvent

# endregion imports


class BuildForm(
    unohelper.Base,
    XEventListener,
    XPropertyChangeListener,
    XActionListener,
    XTextListener,
    XFocusListener,
    XItemListener,
    XMouseListener,
    XSelectionChangeListener,
    XGridControlListener,
):
    doc: XTextDocument = None

    def __init__(self, db_path: Path) -> None:
        super().__init__()

        self._db_fnm = db_path
        self.closed = False
        self._tab_index = 1
        self._init_callbacks()

        loader = Lo.load_office(Lo.ConnectSocket())
        try:
            BuildForm.doc = Write.create_doc(loader)
            self._doc = BuildForm.doc

            GUI.set_visible(True, BuildForm.doc)
            self._top_win_ev = TopWindowEvents(add_window_listener=True)
            self._top_win_ev.add_event_window_closing(self._fn_on_window_closing)
            with Lo.ControllerLock():
                # use a controller lock to lock screen updating.
                # This will cut down and screen flashing and add controls faster.

                BuildForm.doc.addEventListener(self)

                tvc = Write.get_view_cursor(BuildForm.doc)
                Write.append(tvc, "Building a Form\n")
                Write.end_paragraph(tvc)

                self.create_form(BuildForm.doc)
            Lo.dispatch_cmd("SwitchControlDesignMode")

            # Lo.wait_enter()
            # Lo.close_doc(BuildForm.doc)
            # Lo.delay(2000)
            # msg_result = MsgBox.msgbox(
            #     "Do you wish to close document?",
            #     "All done",
            #     boxtype=MessageBoxType.QUERYBOX,
            #     buttons=MessageBoxButtonsEnum.BUTTONS_YES_NO,
            # )
            # if msg_result == MessageBoxResultsEnum.YES:
            #     Lo.close_doc(doc=BuildForm.doc)
            #     BuildForm.doc = None
            #     Lo.close_office()
            # else:
            #     print("Keeping document open")
        except Exception:
            Lo.close_office()
            raise

    def _init_callbacks(self) -> None:
        self._fn_text_changed = self.on_text_changed
        self._fn_on_window_closing = self.on_window_closing
        self._fn_on_up = self.on_up
        self._fn_on_down = self.on_down
        self._fn_on_veto_change_listener = self.on_veto_change_listener
        self._fn_on_property_change_listener = self.on_property_change_listener
        self._fn_FormCtlButton = self.on_btn_action_preformed
        self._fn_on_btn_action_approved = self.on_btn_action_approved
        self._fn_on_mouse_entered = self.on_mouse_entered
        self._fn_on_mouse_exit = self.on_mouse_exit
        self._fn_on_list_item_changed = self.on_list_item_changed
        self._fn_on_grid_column_changed = self.on_grid_column_changed
        self._fn_on_grid_selection_changed = self.on_grid_selection_changed

    def create_form(self, doc: XTextDocument) -> None:
        # Form has four sections: text, command_button, list_box, grid_control
        # Section 1 has two columns
        if Info.version_info < (7, 5, 0, 0):
            dark = False
        else:
            try:
                gen_theme = ThemeGeneral()
                if gen_theme.background_color < 0:
                    # automatic color, assume light
                    dark = False
                else:
                    rgb = color_util.RGB.from_int(gen_theme.background_color)
                    dark = rgb.is_dark()
            except Exception:
                dark = False
        if dark:
            font_color = color_util.StandardColor.WHITE
        else:
            font_color = color_util.StandardColor.BLACK

        _doc = BuildForm.doc

        font_colored = Font(color=font_color)
        x1 = 2
        x2 = 44
        y1 = 11
        height = 6
        width1 = 40
        width2 = 40

        self._ctl_lbl_first_name = Forms.insert_control_label(
            doc,
            label="FIRSTNAME",
            x=x1,
            y=y1,
            width=width1,
            height=height,
            styles=[font_colored],
        )

        self._ctl_txt_first_name = Forms.insert_control_text_field(
            doc,
            x=x2,
            y=y1,
            width=width2,
            height=height,
            border=BorderKind.BORDER_SIMPLE,
        )
        self._ctl_lbl_first_name.bind_to_control(self._ctl_lbl_first_name)
        self._ctl_txt_first_name.add_event_text_changed(self._fn_text_changed)
        self._set_tab_index(self._ctl_txt_first_name)

        # props = Forms.add_labelled_control(
        #     doc=_doc,
        #     label="FIRSTNAME",
        #     comp_kind=FormComponentKind.TEXT_FIELD,
        #     y=11,
        #     lbl_styles=[font_colored],
        # )
        # self.listen_to_text_field(props)

        y1 = 19
        self._ctl_lbl_last_name = Forms.insert_control_label(
            doc,
            label="LASTNAME",
            x=x1,
            y=y1,
            width=width1,
            height=height,
            styles=[font_colored],
        )

        self._ctl_txt_last_name = Forms.insert_control_text_field(
            doc,
            x=x2,
            y=y1,
            width=width2,
            height=height,
            border=BorderKind.BORDER_SIMPLE,
        )
        self._ctl_lbl_last_name.bind_to_control(self._ctl_txt_last_name)
        self._ctl_txt_last_name.add_event_text_changed(self._fn_text_changed)
        self._set_tab_index(self._ctl_txt_last_name)

        # Forms.add_labelled_control(
        #     doc=_doc,
        #     label="LASTNAME",
        #     comp_kind=FormComponentKind.TEXT_FIELD,
        #     y=19,
        #     lbl_styles=[font_colored],
        # )

        y1 = 43

        # props = Forms.add_labelled_control(
        #     doc=_doc, label="AGE", comp_kind=FormComponentKind.NUMERIC_FIELD, y=43, lbl_styles=[font_colored]
        # )
        # Props.set_property(props, "DecimalAccuracy", 0)

        self._ctl_lbl_age = Forms.insert_control_label(
            doc,
            label="AGE",
            x=x1,
            y=y1,
            width=width1,
            height=height,
            styles=[font_colored],
        )

        self._ctl_num_age = Forms.insert_control_numeric_field(
            doc,
            x=x2,
            y=y1,
            width=width2,
            height=height,
            accuracy=0,
            spin_button=False,
            border=BorderKind.BORDER_3D,
        )
        self._ctl_lbl_age.bind_to_control(self._ctl_num_age)
        self._ctl_num_age.add_event_down(self._fn_on_down)
        self._ctl_num_age.add_event_up(self._fn_on_up)
        self._set_tab_index(self._ctl_num_age)

        y1 = 51
        self._ctl_lbl_birth_date = Forms.insert_control_label(
            doc,
            label="BIRTHDATE",
            x=x1,
            y=y1,
            width=width1,
            height=height,
            styles=[font_colored],
        )

        min_date = datetime.datetime.now() - datetime.timedelta(days=365 * 100)
        max_date = datetime.datetime.now()
        self._ctl_date_birth_date = Forms.insert_control_date_field(
            doc,
            x=x2,
            y=y1,
            width=width2,
            height=height,
            min_date=min_date,
            max_date=max_date,
            border=BorderKind.BORDER_3D,
        )
        self._ctl_lbl_birth_date.bind_to_control(self._ctl_date_birth_date)
        self._set_tab_index(self._ctl_date_birth_date)
        # self._ctl_date_birth_date.add_event_text_changed(self._fn_text_changed)

        # Forms.add_labelled_control(
        #     doc=_doc,
        #     label="BIRTHDATE",
        #     comp_kind=FormComponentKind.FORMATTED_FIELD,
        #     y=51,
        #     lbl_styles=[font_colored],
        # )

        # buttons, all with listeners
        col1_x = 2
        x = col1_x
        spacing = 10
        y = 63
        width = 8

        names, labels = map(
            list,
            zip(
                *[
                    ["first", "<<"],
                    ["prev", "<"],
                    ["next", ">"],
                    ["last", ">>"],
                    ["new", ">*"],
                ]
            ),
        )

        # for i in range(0, len(labels)):
        #     props = Forms.add_button(
        #         doc=_doc,
        #         name=names[i],
        #         label=labels[i],
        #         x=x + i * spacing,
        #         y=y,
        #         width=width,
        #     )
        #     self.listen_to_button(props)

        for i in range(len(labels)):
            btn = Forms.insert_control_button(
                doc=doc,
                x=x + i * spacing,
                y=y,
                width=width,
            )
            btn.model.Label = labels[i]
            btn.add_event_action_performed(self._fn_FormCtlButton)
            btn.add_event_approve_action(self._fn_on_btn_action_approved)
            self._set_tab_index(btn)

        self._ctl_btn_reload = Forms.insert_control_button(
            doc=doc, x=x + 4 * spacing + 16, y=y, width=13, name="Button_reload"
        )
        self._ctl_btn_reload.model.Label = "reload"
        self._ctl_btn_reload.add_event_action_performed(self._fn_FormCtlButton)
        self._ctl_btn_reload.add_event_approve_action(self._fn_on_btn_action_approved)
        self._ctl_btn_reload.add_event_mouse_entered(self._fn_on_mouse_entered)
        self._ctl_btn_reload.add_event_mouse_exited(self._fn_on_mouse_exit)
        self._set_tab_index(self._ctl_btn_reload)

        # props = Forms.add_button(
        #     doc=_doc,
        #     name="reload",
        #     label="reload",
        #     x=x + 4 * spacing + 16,
        #     y=y,
        #     width=13,
        # )
        # self.listen_to_button(props)
        # self.listen_to_mouse(props)

        # some fixed text; no listener
        width = 60
        y = 80
        height = 6

        _ = Forms.insert_control_label(
            doc=doc,
            x=x,
            y=y,
            width=width,
            height=height,
            label="show only sales since",
            styles=[font_colored],
        )

        # Forms.add_control(
        #     doc=_doc,
        #     name="text-1",
        #     label="show only sales since",
        #     comp_kind=FormComponentKind.FIXED_TEXT,
        #     x=x,
        #     y=y,
        #     width=width,
        #     height=height,
        #     styles=[font_colored],
        # )

        #  radio buttons inside a group box; use a property change listener
        col2_x = 90
        box_width = 70
        y = 5
        name = "Options"

        _ = Forms.insert_control_group_box(
            doc=doc,
            x=col2_x,
            y=y,
            width=box_width,
            height=25,
            label="Options",
            styles=[font_colored],
        )

        # Forms.add_control(
        #     doc=_doc,
        #     name=name,
        #     label="Options",
        #     comp_kind=FormComponentKind.GROUP_BOX,
        #     x=col2_x,
        #     y=y,
        #     width=box_width,
        #     height=25,
        #     styles=[font_colored],
        # )

        # these three radio buttons have the same name ("Option"), and
        # so only one can be on at a time
        # comp_kind = FormComponentKind.RADIO_BUTTON
        indent = 3
        x = col2_x + indent
        width = box_width - 2 * indent

        labels = [
            "No automatic generation",
            "Before inserting a record",
            "When moving to a new record",
        ]

        for i in range(len(labels)):
            radio_btn = Forms.insert_control_radio_button(
                doc=doc,
                name=name,
                # label=labels[i],
                x=x,
                y=y + (i + 1) * height,
                width=width,
                height=height,
                styles=[font_colored],
            )
            radio_btn.model.Label = labels[i]
            radio_btn.add_event_property_change(
                "State", self._fn_on_property_change_listener
            )
            radio_btn.tag = f"RadioBtn_{i+1}"
            self._set_tab_index(radio_btn)

            # radio_btn.model.addPropertyChangeListener("State", self)

        # for i in range(0, len(labels)):
        #     props = Forms.add_control(
        #         doc=_doc,
        #         name=name,
        #         label=labels[i],
        #         comp_kind=comp_kind,
        #         x=x,
        #         y=y + (i + 1) * height,
        #         width=width,
        #         height=height,
        #         styles=[font_colored],
        #     )
        #     props.addPropertyChangeListener("State", self)

        # check boxes inside another group box
        # use the same property change listener
        y = 35
        height = 25
        x = col2_x
        width = width
        _ = Forms.insert_control_group_box(
            doc=doc,
            x=x,
            y=y,
            width=box_width,
            height=height,
            label="Miscellaneous",
            styles=[font_colored],
        )

        # Forms.add_control(
        #     doc=_doc,
        #     name="Misc",
        #     label="Miscellaneous",
        #     comp_kind=FormComponentKind.GROUP_BOX,
        #     x=col2_x,
        #     y=35,
        #     width=box_width,
        #     height=25,
        #     styles=[font_colored],
        # )

        # comp_kind = FormComponentKind.CHECK_BOX
        y = 33
        width = 60
        x = x + indent
        width = box_width - 4
        height = 6

        # names, labels, HelpTexts = map(
        #     list,
        #     zip(
        #         *[
        #             [
        #                 "DefaultDate",
        #                 'Default sales date to "today"',
        #                 "When checked, newly entered sales records are pre-filled",
        #             ],
        #             [
        #                 "Protect",
        #                 "Protect key fields from editing",
        #                 "When checked, you cannot modify the values",
        #             ],
        #             [
        #                 "Empty",
        #                 "Check for empty sales names",
        #                 "When checked, you cannot enter empty values",
        #             ],
        #         ]
        #     ),
        # )

        # for i in range(0, len(labels)):
        #     props = Forms.add_control(
        #         doc=_doc,
        #         name=names[i],
        #         label=labels[i],
        #         comp_kind=comp_kind,
        #         x=x,
        #         y=y + (i + 1) * height,
        #         width=width,
        #         height=height,
        #         styles=[font_colored],
        #     )
        #     props.addPropertyChangeListener("State", self)
        #     Props.set_property(props, "HelpText", HelpTexts[i])

        y += height
        self._ctl_chk_default_date = Forms.insert_control_check_box(
            doc=doc,
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
        self._ctl_chk_protect = Forms.insert_control_check_box(
            doc=doc,
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
        self._ctl_chk_empty = Forms.insert_control_check_box(
            doc=doc,
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
        y = 90
        x = 2
        self._ctl_lst_fruits = Forms.insert_control_list_box(
            doc=doc,
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
        # props = Forms.add_list(
        #     doc=_doc,
        #     name="Fruits",
        #     entries=fruits,
        #     x=x,
        #     y=y,
        #     width=width,
        #     height=height,
        # )
        # self.listen_to_list(props)

        # set Form's data source to be the DB_FNM database
        def_form = Forms.get_form(doc, "Form")
        Forms.bind_form_to_table(
            xform=def_form, src_name=FileIO.fnm_to_url(self._db_fnm), tbl_name="Course"
        )

        # a list filled using an SQL query on the form's data source
        x = 60

        self._ctl_db_lst_course_names = Forms.insert_db_control_list_box(
            doc=doc,
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

        # props = Forms.add_database_list(
        #     doc=_doc,
        #     name="CourseNames",
        #     sql_cmd='SELECT "title" FROM "Course"',
        #     x=x,
        #     y=y,
        #     width=width,
        #     height=height,
        # )
        # self.listen_to_list(props)

        # another list filled using a different SQL query on the form's data source
        x = 120
        # ------------------------ set up database grid/table ----------------
        self._ctl_db_lst_stud_names = Forms.insert_db_control_list_box(
            doc=doc,
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

        # props = Forms.add_database_list(
        #     doc=_doc,
        #     name="StudNames",
        #     sql_cmd='SELECT "lastName" FROM "Student"',
        #     x=x,
        #     y=y,
        #     width=width,
        #     height=height,
        # )
        # self.listen_to_list(props)

        # create a new form, gridForm,
        grid_con = Forms.insert_form(doc=_doc)
        grid_form = Lo.qi(XForm, grid_con, True)

        # which uses an SQL query as its data source
        Forms.bind_form_to_sql(
            xform=grid_form,
            src_name=FileIO.fnm_to_url(self._db_fnm),
            cmd='SELECT "firstName", "lastName" FROM "Student"',
        )

        # create the grid/table component and set its columns
        self._ctl_grid_sales_tbl = Forms.insert_control_grid(
            doc=doc,
            name="GridSalesTable",
            x=2,
            y=100,
            width=100,
            height=40,
            parent_form=grid_con,
        )
        self._ctl_grid_sales_tbl.create_grid_column(
            data_field="firstName", col_kind="TextField", width=25
        )
        self._ctl_grid_sales_tbl.create_grid_column(
            data_field="lastName", col_kind="TextField", width=25
        )
        self._set_tab_index(self._ctl_grid_sales_tbl)
        self._ctl_grid_sales_tbl.add_event_selection_changed(self._fn_on_grid_selection_changed)
        self._ctl_grid_sales_tbl.add_event_column_changed(self._fn_on_grid_column_changed)
        #  _fn_on_grid_selection_changed

        # props = Forms.add_control(
        #     doc=_doc,
        #     name="SalesTable",
        #     label=None,
        #     comp_kind=FormComponentKind.GRID_CONTROL,
        #     x=2,
        #     y=100,
        #     width=100,
        #     height=40,
        #     parent_form=grid_con,
        # )

        # grid_model = Lo.qi(XControlModel, props, True)
        # Forms.create_grid_column(
        #     grid_model=grid_model,
        #     data_field="firstName",
        #     col_kind="TextField",
        #     width=25,
        # )
        # Forms.create_grid_column(
        #     grid_model=grid_model, data_field="lastName", col_kind="TextField", width=25
        # )

        # self.listen_to_gird(grid_model)

    def _set_tab_index(self, ctl: DialogControlBase) -> None:
        ctl.tab_index = self._tab_index
        self._tab_index += 1

    def listen_to_button(self, props: XPropertySet) -> None:
        cmodel = Lo.qi(XControlModel, props, True)
        control = Forms.get_control(doc=BuildForm.doc, ctl_model=cmodel)
        xbutton = Lo.qi(XButton, control)

        xbutton.setActionCommand(Forms.get_name(cmodel))
        xbutton.addActionListener(self)

    def listen_to_text_field(self, props: XPropertySet) -> None:
        cmodel = Lo.qi(XControlModel, props, True)
        control = Forms.get_control(doc=BuildForm.doc, ctl_model=cmodel)
        # convert the control into two components, so two different
        # listeners can be attached
        tc = Lo.qi(XTextComponent, control, True)
        tc.addTextListener(self)

        tf_window = Lo.qi(XWindow, control, True)
        tf_window.addFocusListener(self)

    # region XTextListener
    def on_text_changed(
        self, src: Any, event: EventArgs, control_src: FormCtlTextField, *args, **kwargs
    ) -> None:
        # ev_arg = cast("TextEvent", event.event_data)
        print(f"listen_to_text_field2: {control_src.name}: {control_src.text}")

    # endregion XTextListener

    def listen_to_list(self, props: XPropertySet) -> None:
        cmodel = Lo.qi(XControlModel, props, True)
        control = Forms.get_control(doc=BuildForm.doc, ctl_model=cmodel)

        list_box = Lo.qi(XListBox, control)
        list_box.addItemListener(self)

    def listen_to_gird(self, grid_model: XControlModel) -> None:
        control = Forms.get_control(doc=BuildForm.doc, ctl_model=grid_model)
        gc = Lo.qi(XGridControl, control, True)
        gc.addGridControlListener(self)

        grid_selection = Lo.qi(XSelectionSupplier, gc, True)
        grid_selection.addSelectionChangeListener(self)

    def listen_to_mouse(self, props: XPropertySet) -> None:
        cmodel = Lo.qi(XControlModel, props, True)
        control = Forms.get_control(doc=BuildForm.doc, ctl_model=cmodel)
        xwindow = Lo.qi(XWindow, control)
        xwindow.addMouseListener(self)

    # region XEventListener
    def disposing(self, ev: EventObject) -> None:
        imp_name = Info.get_implementation_name(ev.Source)
        print(f"Disposing: {imp_name}")

    # endregion XEventListener

    # region XActionListener
    def actionPerformed(self, ev: ActionEvent) -> None:
        print(f'Pressed "{ev.ActionCommand}"')

    # endregion XActionListener

    # region XTextListener
    def textChanged(self, ev: TextEvent) -> None:
        cmodel = Forms.get_event_control_model(ev)
        print(
            f"{Forms.get_event_control_model(ev)}, Text: {Props.get_property(cmodel, 'Text')}"
        )

    # endregion XTextListener

    # region XFocusListener
    # also used by the text field
    def focusGained(self, ev: FocusEvent) -> None:
        cmodel = Forms.get_event_control_model(ev)
        print(f"Into: {Forms.get_name(cmodel)}")

    def focusLost(self, ev: FocusEvent) -> None:
        cmodel = Forms.get_event_control_model(ev)
        print(f"Left: {Forms.get_name(cmodel)}")

    # endregion XFocusListener

    # region XPropertyChangeListener
    def propertyChange(self, ev: PropertyChangeEvent) -> None:
        print("Property Change detected")
        name = str(Props.get_property(ev.Source, "Name"))
        lbl = str(Props.get_property(ev.Source, "Label"))
        cls_id = int(Props.get_property(ev.Source, "ClassId"))
        is_enabled = int(ev.NewValue) != 0

        # did it come from a radio button or checkbox?
        if cls_id == FormComponentType.RADIOBUTTON:
            # use the label since all my radio buttons have the same name
            print(f'"{lbl}" radio button: {is_enabled}')
        elif cls_id == FormComponentType.CHECKBOX:
            print(f'"{name}" checkbox: {is_enabled}')

    # endregion XPropertyChangeListener

    # region XItemListener
    def itemStateChanged(self, ev: ItemEvent) -> None:
        selection = int(ev.Selected)
        selected_item = None

        cmodel = Forms.get_event_control_model(ev)
        list_box = Lo.qi(XListBox, ev.Source)
        if list_box is not None:
            selected_item = list_box.getItem(selection)
        print(f'List: "{Forms.get_name(cmodel)}" selection: {selected_item}')

    # endregion XItemListener

    # region XMouseListener
    def mousePressed(self, ev: MouseEvent) -> None:
        cmodel = Forms.get_event_control_model(ev)
        print(f"Pressed {Forms.get_name(cmodel)} at ({ev.X}, {ev.Y})")

    def mouseReleased(self, ev: MouseEvent) -> None:
        pass

    def mouseEntered(self, ev: MouseEvent) -> None:
        cmodel = Forms.get_event_control_model(ev)
        print(f"Entered {Forms.get_name(cmodel)}")

    def mouseExited(self, ev: MouseEvent) -> None:
        pass

    # endregion XMouseListener

    # region XSelectionChangeListener
    def selectionChanged(self, ev: EventObject) -> None:
        # used by the grid/table control
        cmodel = Forms.get_event_control_model(ev)
        gc = Lo.qi(XGridControl, ev.Source)

        # in the form module
        # not the one in com.sun.star.awt.grid
        # no way to get the current row with XGridControl; see form.XGrid
        print(f"Grid {Forms.get_name(cmodel)} column: {gc.getCurrentColumnPosition()}")

        # must access the result set inside the form for row info
        form_name = Forms.get_form_name(cmodel)
        gform = Forms.get_form(BuildForm.doc, form_name)
        rs = Lo.qi(XResultSet, gform)
        try:
            print(f"    row: {rs.getRow()}")
        except Exception as e:
            print(e)

    # endregion XSelectionChangeListener

    # region XGridControlListener
    def columnChanged(self, ev: EventObject) -> None:
        # used by the grid/table control
        print("Grid Column change")
    # endregion XGridControlListener

    def on_window_closing(
        self, source: Any, event_args: EventArgs, *args, **kwargs
    ) -> None:
        print("Closing")
        try:
            Lo.close_doc(self._doc)
            Lo.close_office()
            self.closed = True
        except Exception as e:
            print(f"  {e}")


    def on_up(
        self, src: Any, event: EventArgs, control_src: Any, *args, **kwargs
    ) -> None:
        print("Up:", control_src.name)
        print("Value:", control_src.value)

    def on_down(
        self, src: Any, event: EventArgs, control_src: Any, *args, **kwargs
    ) -> None:
        print("Down:", control_src.name)
        print("Value:", control_src.value)

    def on_veto_change_listener(
        self, src: Any, event: EventArgs, control_src: Any, *args, **kwargs
    ) -> None:
        ev_data = cast("PropertyChangeEvent", event.event_data)
        print("Vetoable Change:", control_src.name)
        print("Old Value:", ev_data.OldValue)
        print("New Value:", ev_data.NewValue)

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

    def on_btn_action_preformed(
        self, src: Any, event: EventArgs, control_src: FormCtlButton, *args, **kwargs
    ) -> None:
        print(
            f"Action Preformed: '{control_src.model.Label}', Control Name: {control_src.name}"
        )

    def on_btn_action_approved(
        self, src: Any, event: EventArgs, control_src: FormCtlButton, *args, **kwargs
    ) -> None:
        print(
            f"Action Approved: '{control_src.model.Label}', Control Name: {control_src.name}"
        )

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

    def on_list_item_changed(
        self, src: Any, event: EventArgs, control_src: FormCtlListBox, *args, **kwargs
    ) -> None:
        itm_event = cast("ItemEvent", event.event_data)
        # print("Selected:", itm_event.Selected)
        # print("Highlighted:", itm_event.Highlighted)
        selected_item = control_src.get_item(itm_event.Selected)
        print(f'List: "{control_src.name}", selection: {selected_item}')
    
    
    def on_grid_column_changed(
        self, src: Any, event: EventArgs, control_src: FormCtlGrid, *args, **kwargs
    ) -> None:
        # print(control_src)
        print("Grid Column change", control_src.name)

    def on_grid_selection_changed(self, src: Any, event: EventArgs, control_src: FormCtlGrid, *args, **kwargs) -> None:
        # used by the grid/table control

        # in the form module
        # not the one in com.sun.star.awt.grid
        # no way to get the current row with XGridControl; see form.XGrid
        print(f"Grid {control_src.name} column: {control_src.get_current_column_position()}")

        # must access the result set inside the form for row info
        form_name = control_src.get_form_name()
        if not form_name:
            return
        gform = Forms.get_form(self._doc, form_name)
        rs = Lo.qi(XResultSet, gform)
        if not rs:
            return
        try:
            print(f"    row: {rs.getRow()}")
        except Exception as e:
            print(e)