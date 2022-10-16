# region imports
from __future__ import annotations
from pathlib import Path

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
from com.sun.star.form import XGridControl
from com.sun.star.form import XGridControlListener
from com.sun.star.lang import EventObject
from com.sun.star.sdbc import XResultSet
from com.sun.star.text import XTextDocument
from com.sun.star.view import XSelectionChangeListener
from com.sun.star.view import XSelectionSupplier
from com.sun.star.form import XForm
import unohelper

from ooo.dyn.form.form_component_type import FormComponentType

from ooodev.office.write import Write
from ooodev.utils.forms import Forms
from ooodev.utils.gui import GUI
from ooodev.utils.info import Info
from ooodev.utils.lo import Lo
from ooodev.utils.props import Props
from ooodev.utils.file_io import FileIO

# from .....resources import __res_path__
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

        loader = Lo.load_office(Lo.ConnectSocket())
        BuildForm.doc = Write.create_doc(loader)

        GUI.set_visible(True, BuildForm.doc)
        with Lo.ControllerLock():
            # use a controller lock to lock screen updating.
            # This will cut down and screen flashiing and add controls faster.
            BuildForm.doc.addEventListener(self)

            tvc = Write.get_view_cursor(BuildForm.doc)
            Write.append(tvc, "Building a Form\n")
            Write.end_paragraph(tvc)

            self.create_form(BuildForm.doc)
        Lo.dispatch_cmd("SwitchControlDesignMode")

        Lo.wait_enter()
        Lo.close_doc(BuildForm.doc)
        Lo.close_office()

    def create_form(self, doc: XTextDocument) -> None:

        # Form has four sections: text, command_button, list_box, grid_control
        # Section 1 has two columns
        _doc = BuildForm.doc

        props = Forms.add_labelled_control(doc=_doc, label="FIRSTNAME", comp_kind=Forms.CompenentKind.TextField, y=11)
        self.listen_to_text_field(props)

        Forms.add_labelled_control(doc=_doc, label="LASTNAME", comp_kind=Forms.CompenentKind.TextField, y=19)

        props = Forms.add_labelled_control(doc=_doc, label="AGE", comp_kind=Forms.CompenentKind.NumericField, y=43)
        Props.set_property(props, "DecimalAccuracy", 0)

        Forms.add_labelled_control(
            doc=_doc,
            label="BIRTHDATE",
            comp_kind=Forms.CompenentKind.FormattedField,
            y=51,
        )

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

        for i in range(0, len(labels)):
            props = Forms.add_button(
                doc=_doc,
                name=names[i],
                label=labels[i],
                x=x + i * spacing,
                y=y,
                width=width,
            )
            self.listen_to_button(props)

        props = Forms.add_button(
            doc=_doc,
            name="reload",
            label="reload",
            x=x + 4 * spacing + 16,
            y=y,
            width=13,
        )
        self.listen_to_button(props)
        self.listen_to_mouse(props)

        # some fixed text; no listener
        width = 60
        y = 80
        height = 6
        Forms.add_control(
            doc=_doc,
            name="text-1",
            label="show only sales since",
            comp_kind=Forms.CompenentKind.FixedText,
            x=x,
            y=y,
            width=width,
            height=height,
        )

        #  radio buttons inside a group box; use a property change listener
        col2_x = 90
        col2_width = 70
        y = 5

        name = "Options"

        Forms.add_control(
            doc=_doc,
            name=name,
            label="Options",
            comp_kind=Forms.CompenentKind.GroupBox,
            x=col2_x,
            y=y,
            width=col2_width,
            height=25,
        )

        # these three radio buttons have the same name ("Option"), and
        # so only one can be on at a time
        comp_kind = Forms.CompenentKind.RadioButton
        indent = 3
        x = col2_x + indent
        width = col2_width - 2 * indent

        labels = [
            "No automatic generation",
            "Before inserting a record",
            "When moving to a new record",
        ]

        for i in range(0, len(labels)):
            props = Forms.add_control(
                doc=_doc,
                name=name,
                label=labels[i],
                comp_kind=comp_kind,
                x=x,
                y=y + (i + 1) * height,
                width=width,
                height=height,
            )
            props.addPropertyChangeListener("State", self)

        # check boxes inside another group box
        # use the same property change listener
        y = 33
        width = 60

        x = col2_x
        width = width
        Forms.add_control(
            doc=_doc,
            name="Misc",
            label="Miscellaneous",
            comp_kind=Forms.CompenentKind.GroupBox,
            x=col2_x,
            y=35,
            width=col2_width,
            height=25,
        )

        comp_kind = Forms.CompenentKind.CheckBox
        x = x + indent
        width = col2_width - 4

        names, labels, HelpTexts = map(
            list,
            zip(
                *[
                    [
                        "DefaultDate",
                        'Default sales date to "today"',
                        "When checked, newly entered sales records are pre-filled",
                    ],
                    [
                        "Protect",
                        "Protect key fields from editing",
                        "When checked, you cannot modify the values",
                    ],
                    [
                        "Empty",
                        '"Check for empty sales names"',
                        "When checked, you cannot enter empty values",
                    ],
                ]
            ),
        )

        for i in range(0, len(labels)):
            props = Forms.add_control(
                doc=_doc,
                name=names[i],
                label=labels[i],
                comp_kind=comp_kind,
                x=x,
                y=y + (i + 1) * height,
                width=width,
                height=height,
            )
            props.addPropertyChangeListener("State", self)
            Props.set_property(props, "HelpText", HelpTexts[i])
            props.addPropertyChangeListener("State", self)

        # a list using simple text
        fruits = ("apple", "orange", "pear", "grape")
        width = 40
        height = 6
        y = 90
        x = 2
        props = Forms.add_list(doc=_doc, name="Fruits", entries=fruits, x=x, y=y, width=width, height=height)
        self.listen_to_list(props)

        # set Form's data source to be the DB_FNM database
        def_form = Forms.get_form(doc, "Form")
        Forms.bind_form_to_table(xform=def_form, src_name=FileIO.fnm_to_url(self._db_fnm), tbl_name="Course")

        # a list filled using an SQL query on the form's data source
        x = 60
        props = Forms.add_database_list(
            doc=_doc,
            name="CourseNames",
            sql_cmd='SELECT "title" FROM "Course"',
            x=x,
            y=y,
            width=width,
            height=height,
        )
        self.listen_to_list(props)

        # another list filled using a different SQL query on the form's data source
        x = 120
        # ------------------------ set up database grid/table ----------------
        props = Forms.add_database_list(
            doc=_doc,
            name="StudNames",
            sql_cmd='SELECT "lastName" FROM "Student"',
            x=x,
            y=y,
            width=width,
            height=height,
        )
        self.listen_to_list(props)

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
        props = Forms.add_control(
            doc=_doc,
            name="SalesTable",
            label=None,
            comp_kind=Forms.CompenentKind.GridControl,
            x=2,
            y=100,
            width=100,
            height=40,
            parent_form=grid_con,
        )

        grid_model = Lo.qi(XControlModel, props, True)
        Forms.create_grid_column(
            grid_model=grid_model,
            data_field="firstName",
            col_kind="TextField",
            width=25,
        )
        Forms.create_grid_column(grid_model=grid_model, data_field="lastName", col_kind="TextField", width=25)

        self.listen_to_gird(grid_model)

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
        print(f"{Forms.get_event_control_model(ev)}, Text: {Props.get_property(cmodel, 'Text')}")

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
        if not list_box is None:
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
