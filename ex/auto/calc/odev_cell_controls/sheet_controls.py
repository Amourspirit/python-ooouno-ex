from __future__ import annotations
import uno
from com.sun.star.awt import XActionListener
from typing import TYPE_CHECKING
from ooodev.calc import CalcDoc
from ooodev.utils.kind.language_kind import LanguageKind
from ooodev.utils.color import StandardColor
from ooodev.utils.data_type.range_obj import RangeObj

if TYPE_CHECKING:
    from ooodev.form.controls import FormCtlButton


class SheetControls:
    def __init__(self, doc: CalcDoc):
        self._doc = doc
        self._script_name = "odev_cell_controls.py"
        self._start_row = 2
        self._end_row = 4

        self._add_info_buttons()
        self._add_button_auto_ctl_start()
        self._add_button_auto_ctl_stop()
        # self._add_currency()
        self._doc.toggle_design_mode()
        sheet = self._doc.sheets[0]
        sheet.protect_sheet(password="")

    def _assign_btn_script(self, ctl: FormCtlButton, script_loc: str) -> None:
        ctl.assign_script(
            interface_name=XActionListener,
            method_name="actionPerformed",
            script_name=script_loc,
            loc="document",
            language=LanguageKind.PYTHON,
        )

    def _add_info_buttons(self) -> None:
        sheet = self._doc.sheets[0]
        used_range = sheet.find_used_range_obj()
        # https://python-ooo-dev-tools.readthedocs.io/en/latest/help/common/ranges/range_obj.html
        # drop the first row and get only the E column
        nxt_col_name = str((used_range.cell_end + "A").col)
        col_rng = RangeObj(
            col_start=nxt_col_name,
            col_end=nxt_col_name,
            row_start=2,
            row_end=used_range.row_end,
            sheet_idx=sheet.sheet_index,
        )

        script_loc = f"{self._script_name}$on_btn_action_preformed"
        for i, cell_obj in enumerate(col_rng):
            cell = sheet[cell_obj]
            ctl = cell.control.insert_control_button(f"Info Row: {i+2}")
            self._assign_btn_script(ctl, script_loc)

    def _add_button_auto_ctl_start(self) -> None:
        sheet = self._doc.sheets[0]
        # work with the first 10 rows
        cell = sheet["G2"]
        sheet.set_col_width(40, cell.cell_obj.col_obj.index)
        script_loc = f"{self._script_name}$start_auto_control"
        ctl = cell.control.insert_control_button("Start Auto Controls")
        ctl.model.BackgroundColor = StandardColor.GREEN_LIGHT2
        ctl.help_text = "Starts the auto control of the cells"
        self._assign_btn_script(ctl, script_loc)

    def _add_button_auto_ctl_stop(self) -> None:
        sheet = self._doc.sheets[0]
        # work with the first 10 rows
        cell = sheet["H2"]
        sheet.set_col_width(40, cell.cell_obj.col_obj.index)
        script_loc = f"{self._script_name}$stop_auto_control"
        ctl = cell.control.insert_control_button("Stop Auto Controls")
        ctl.help_text = "Stops the auto control of the cells"
        ctl.model.BackgroundColor = StandardColor.RED_LIGHT1
        self._assign_btn_script(ctl, script_loc)

    def _add_currency(self) -> None:
        sheet = self._doc.sheets[0]
        # work with the first 10 rows
        rng = sheet.rng(f"B{self._start_row}:B{self._end_row}")
        script_loc = f"{self._script_name}$on_text_changed"

        for i, cell_obj in enumerate(rng):
            cell = sheet[cell_obj]
            ctl = cell.control.insert_control_currency_field(
                min_value=0, spin_button=True
            )
            ctl.value = cell.get_num()
            ctl.assign_script(
                interface_name="com.sun.star.awt.XTextListener",
                method_name="textChanged",
                script_name=script_loc,
                loc="document",
                language=LanguageKind.PYTHON,
            )
