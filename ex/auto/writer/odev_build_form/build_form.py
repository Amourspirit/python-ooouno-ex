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
        props = Forms.add_labelled_control(doc=doc, label="FIRSTNAME", comp_kind=Forms.CompenentKind.TextField, y=11)
        self.listen_to_text_field(props)

        Forms.add_labelled_control(doc=BuildForm.doc, label="LASTNAME", comp_kind=Forms.CompenentKind.TextField, y=19)

        props = Forms.add_labelled_control(
            doc=BuildForm.doc, label="AGE", comp_kind=Forms.CompenentKind.NumericField, y=43
        )
        Props.set_property(props, "DecimalAccuracy", 0)

        Forms.add_labelled_control(
            doc=BuildForm.doc, label="BIRTHDATE", comp_kind=Forms.CompenentKind.FormattedField, y=51
        )

        # buttons, all with listeners
        props = Forms.add_button(doc=BuildForm.doc, name="first", label="<<", x=2, y=63, width=8)
        self.listen_to_button(props)

        props = Forms.add_button(doc=BuildForm.doc, name="prev", label="<", x=12, y=63, width=8)
        self.listen_to_button(props)

        props = Forms.add_button(doc=BuildForm.doc, name="next", label=">", x=22, y=63, width=8)
        self.listen_to_button(props)

        props = Forms.add_button(doc=BuildForm.doc, name="last", label=">>", x=32, y=63, width=8)
        self.listen_to_button(props)

        props = Forms.add_button(doc=BuildForm.doc, name="new", label=">*", x=42, y=63, width=8)
        self.listen_to_button(props)

        props = Forms.add_button(doc=BuildForm.doc, name="reload", label="reload", x=58, y=63, width=13)
        self.listen_to_button(props)
        self.listen_to_mouse(props)

        # some fixed text; no listener
        Forms.add_control(
            doc=BuildForm.doc,
            name="text-1",
            label="show only sales since",
            comp_kind=Forms.CompenentKind.FixedText,
            x=2,
            y=80,
            width=35,
            height=6,
        )

        #  radio buttons inside a group box; use a property change listener
        Forms.add_control(
            doc=BuildForm.doc,
            name="Options",
            label="Options",
            comp_kind=Forms.CompenentKind.GroupBox,
            x=103,
            y=5,
            width=56,
            height=25,
        )

        # these three radio buttons have the same name ("Option"), and
        # so only one can be on at a time
        props = Forms.add_control(
            doc=BuildForm.doc,
            name="Options",
            label="No automatic generation",
            comp_kind=Forms.CompenentKind.RadioButton,
            x=106,
            y=11,
            width=50,
            height=6,
        )
        props.addPropertyChangeListener("State", self)

        props = Forms.add_control(
            doc=BuildForm.doc,
            name="Options",
            label="Before inserting a record",
            comp_kind=Forms.CompenentKind.RadioButton,
            x=106,
            y=17,
            width=50,
            height=6,
        )
        props.addPropertyChangeListener("State", self)

        props = Forms.add_control(
            doc=BuildForm.doc,
            name="Options",
            label="When moving to a new record",
            comp_kind=Forms.CompenentKind.RadioButton,
            x=106,
            y=23,
            width=50,
            height=6,
        )
        props.addPropertyChangeListener("State", self)

        # check boxes inside another group box
        # use the same property change listener
        Forms.add_control(
            doc=BuildForm.doc,
            name="Misc",
            label="Miscellaneous",
            comp_kind=Forms.CompenentKind.GroupBox,
            x=103,
            y=35,
            width=56,
            height=25,
        )

        props = Forms.add_control(
            doc=BuildForm.doc,
            name="DefaultDate",
            label='Default sales date to "today"',
            comp_kind=Forms.CompenentKind.CheckBox,
            x=106,
            y=39,
            width=60,
            height=6,
        )
        Props.set_property(props, "HelpText", "When checked, newly entered sales records are pre-filled")
        props.addPropertyChangeListener("State", self)

        props = Forms.add_control(
            doc=BuildForm.doc,
            name="Protect",
            label="Protect key fields from editing",
            comp_kind=Forms.CompenentKind.CheckBox,
            x=106,
            y=45,
            width=60,
            height=6,
        )
        Props.set_property(props, "HelpText", "When checked, you cannot modify the values")
        props.addPropertyChangeListener("State", self)

        props = Forms.add_control(
            doc=BuildForm.doc,
            name="Empty",
            label="Check for empty sales names",
            comp_kind=Forms.CompenentKind.CheckBox,
            x=106,
            y=51,
            width=60,
            height=6,
        )
        Props.set_property(props, "HelpText", "When checked, you cannot enter empty values")
        props.addPropertyChangeListener("State", self)

        # a list using simple text
        fruits = ("apple", "orange", "pear", "grape")
        props = Forms.add_list(doc=BuildForm.doc, name="Fruits", entries=fruits, x=2, y=90, width=20, height=6)
        self.listen_to_list(props)

        # set Form's data source to be the DB_FNM database
        def_form = Forms.get_form(doc, "Form")
        Forms.bind_form_to_table(xform=def_form, src_name=FileIO.fnm_to_url(self._db_fnm), tbl_name="Course")

        # a list filled using an SQL query on the form's data source
        props = Forms.add_database_list(
            doc=BuildForm.doc,
            name="CourseNames",
            sql_cmd='SELECT "title" FROM "Course"',
            x=90,
            y=90,
            width=20,
            height=6,
        )
        self.listen_to_list(props)

        # another list filled using a different SQL query on the form's data source

        # ------------------------ set up database grid/table ----------------
        props = Forms.add_database_list(
            doc=BuildForm.doc,
            name="StudNames",
            sql_cmd='SELECT "lastName" FROM "Student"',
            x=140,
            y=90,
            width=20,
            height=6,
        )
        self.listen_to_list(props)
        
        # create a new form, gridForm,
        grid_con = Forms.insert_form(doc=BuildForm.doc)
        grid_form = Lo.qi(XForm, grid_con, True)

        # which uses an SQL query as its data source
        Forms.bind_form_to_sql(
            xform=grid_form,
            src_name=FileIO.fnm_to_url(self._db_fnm),
            cmd='SELECT "firstName", "lastName" FROM "Student"',
        )

        # create the grid/table component and set its columns
        props = Forms.add_control(
            doc=BuildForm.doc,
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
        Forms.create_grid_column(grid_model=grid_model, data_field="firstName", col_kind="TextField", width=25)
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
