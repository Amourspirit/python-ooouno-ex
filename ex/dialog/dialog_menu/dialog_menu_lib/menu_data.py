from __future__ import annotations
from typing import List
from pathlib import Path
import uno
from ooo.dyn.awt.menu_item_style import MenuItemStyleEnum
from ooodev.gui.menu.popup_menu import PopupMenu
from ooodev.gui.menu.popup.popup_creator import PopupCreator


def get_popup_menu() -> PopupMenu:
    menu_data = get_menu_data()
    popup_creator = PopupCreator()
    popup_menu = popup_creator.create(menu_data)
    return popup_menu


def get_popup_from_json(fnm: str | Path) -> PopupMenu:
    """Load a menu from Json data"""
    # https://tinyurl.com/2bh5tj6h

    menus = PopupCreator.json_load(fnm)
    creator = PopupCreator()
    return creator.create(menus)


def get_menu_data() -> List[dict]:
    new_menu = [
        {
            "text": "File",
            "command": "file",
            "submenu": [
                {"text": "Close OK", "command": ".uno:exitok"},
                {"text": "Close", "command": ".uno:exit"},
            ],
        },
        {
            "text": "Edit",
            "command": ".uno:EditMenu",
            "submenu": [
                {"text": "Cut", "command": ".uno:Cut", "shortcut": "Ctrl+X"},
                {"text": "Copy", "command": ".uno:Copy", "shortcut": "Ctrl+C"},
                {"text": "Paste", "command": ".uno:Paste", "shortcut": "Ctrl+V"},
                {
                    "text": "Paste Special",
                    "command": ".uno:PasteSpecialMenu",
                    "submenu": [
                        {
                            "text": "Paste Unformatted",
                            "command": ".uno:PasteUnformatted",
                        },
                        {"text": "-"},
                        {
                            "text": "My Paste Only Text",
                            "command": ".uno:PasteOnlyText",
                        },
                        {"text": "Paste Only Text", "command": ".uno:PasteOnlyValue"},
                        {
                            "text": "Paste Only Formula",
                            "command": ".uno:PasteOnlyFormula",
                        },
                        {"text": "-"},
                        {"text": "Paste Transposed", "command": ".uno:PasteTransposed"},
                        {"text": "-"},
                        {
                            "text": "Paste Special ...",
                            "command": ".uno:PasteSpecial",
                        },
                    ],
                },
                {"text": "-"},
                {"text": "Data Select", "command": ".uno:DataSelect"},
                {"text": "Current Validation", "command": ".uno:CurrentValidation"},
                {"text": "Define Current Name", "command": ".uno:DefineCurrentName"},
                {"text": "-"},
                {"text": "Insert cells", "command": ".uno:InsertCell"},
                {"text": "Del cells", "command": ".uno:DeleteCell"},
                {"text": "Delete", "command": ".uno:Delete"},
                {"text": "Merge Cells", "command": ".uno:MergeCells"},
                {"text": "Split Cell", "command": ".uno:SplitCell"},
                {"text": "-"},
                {"text": "Format Paintbrush", "command": ".uno:FormatPaintbrush"},
                {"text": "Reset Attributes", "command": ".uno:ResetAttributes"},
                {
                    "text": "Format Styles Menu",
                    "command": ".uno:FormatStylesMenu",
                    "submenu": [
                        {"text": "Edit Style", "command": ".uno:EditStyle"},
                        {"text": "-"},
                        {
                            "text": "Default Cell Styles",
                            "command": ".uno:DefaultCellStylesmenu",
                            "style": MenuItemStyleEnum.RADIOCHECK,
                        },
                        {
                            "text": "Accent1 Cell Styles",
                            "command": ".uno:Accent1CellStyles",
                            "style": MenuItemStyleEnum.RADIOCHECK,
                        },
                        {
                            "text": "Accent2 Cell Styles",
                            "style": MenuItemStyleEnum.RADIOCHECK,
                        },
                        {
                            "text": "Accent 3 Cell Styles",
                            "command": ".uno:Accent3CellStyles",
                            "style": MenuItemStyleEnum.RADIOCHECK,
                        },
                        {"text": "-"},
                        {
                            "text": "Bad Cell Styles",
                            "command": ".uno:BadCellStyles",
                            "style": MenuItemStyleEnum.RADIOCHECK,
                        },
                        {
                            "text": "Error Cell Styles",
                            "command": ".uno:ErrorCellStyles",
                            "style": MenuItemStyleEnum.RADIOCHECK,
                        },
                        {
                            "text": "Good Cell Styles",
                            "command": ".uno:GoodCellStyles",
                            "style": MenuItemStyleEnum.RADIOCHECK,
                        },
                        {
                            "text": "Neutral Cell Styles",
                            "command": ".uno:NeutralCellStyles",
                            "style": MenuItemStyleEnum.RADIOCHECK,
                        },
                        {
                            "text": "Warning Cell Styles",
                            "command": ".uno:WarningCellStyles",
                            "style": MenuItemStyleEnum.RADIOCHECK,
                        },
                        {
                            "text": "-",
                        },
                        {
                            "text": "Footnote Cell Styles",
                            "command": ".uno:FootnoteCellStyles",
                            "style": MenuItemStyleEnum.RADIOCHECK,
                        },
                        {
                            "text": "Note Cell Styles",
                            "command": ".uno:NoteCellStyles",
                            "style": MenuItemStyleEnum.RADIOCHECK,
                        },
                    ],
                },
                {"text": "-"},
                {"text": "Insert Annotation", "command": ".uno:InsertAnnotation"},
                {"text": "Edit Annotation", "command": ".uno:EditAnnotation"},
                {"text": "Delete Note", "command": ".uno:DeleteNote"},
                {"text": "Show Note", "command": ".uno:ShowNote"},
                {"text": "Hide Note", "command": ".uno:HideNote"},
                {"text": "-"},
                {"text": "Format Sparkline", "command": ".uno:FormatSparklineMenu"},
                {"text": "-"},
                {
                    "text": "Current Conditional Format Dialog ...",
                    "command": ".uno:CurrentConditionalFormatDialog",
                },
                {
                    "text": "Current Conditional Format Manager Dialog ...",
                    "command": ".uno:CurrentConditionalFormatManagerDialog",
                },
                {"text": "Format Cell Dialog ...", "command": ".uno:FormatCellDialog"},
            ],
        },
        {
            "text": "Help",
            "command": ".uno:HelpMenu",
            "submenu": [
                {"text": "Help", "command": ".uno:HelpIndex"},
                {"text": "What's This?", "command": ".uno:WhatsThis"},
                {"text": "-"},
                {"text": "About", "command": ".uno:About"},
            ],
        },
    ]
    return new_menu
