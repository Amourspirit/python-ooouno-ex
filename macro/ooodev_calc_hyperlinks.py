from __future__ import annotations
from ooodev.calc import CalcDoc, CalcSheet
from ooodev.dialog.msgbox import MsgBox

# This simple macro will write a list of hyperlinks as text into a Calc sheet.
# The links are then converted to hyperlinks.

def set_cell_data(sheet: CalcSheet) -> None:
    vals = (
        ("Hyperlinks",),
        ("https://ask.libreoffice.org/t/how-to-convert-links-into-hyperlinks-in-bulk-in-calc/102448",),
        ("https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/calc/odev_garlic_secrets",),
        ("https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/calc/odev_cell_texts",),
        ("https://python-ooo-dev-tools.readthedocs.io/en/latest/src/utils/data_type/range_obj.html",),
        ("https://python-ooo-dev-tools.readthedocs.io/en/latest/index.html",),
        ("https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part4/index.html",),
    )
    sheet.set_array(values=vals, name="A1")
    # bold the header
    _ = sheet["A1"].style_font_general(b=True)
    col = sheet.get_col_range(0)
    col.optimal_width = True


def create_hyperlinks(doc: CalcDoc) -> None:
    # get access to current Calc Document

    # get access to first spreadsheet
    sheet = doc.get_active_sheet()

    # insert the array of data
    set_cell_data(sheet=sheet)
    convert_to_hyperlinks(sheet=sheet)


def convert_to_hyperlinks(sheet: CalcSheet) -> None:
    # convert the text to hyperlinks
    used_rng = sheet.find_used_range_obj()
    data = sheet.get_array(range_obj=used_rng)
    row_count = 0
    for row in data:
        for i, cell_data in enumerate(row):
            if cell_data.startswith("http"):
                cell = sheet[(i, row_count)]
                cell.value = ""
                cursor = cell.create_text_cursor()
                cursor.add_hyperlink(
                    label=cell_data,
                    url_str=cell_data,
                )
        row_count += 1


def make_hyperlinks(*args) -> None:
    """
    Writes links as text into a sheet and then converts them to hyperlinks.

    If not a Calc document then a error message is displayed.

    Returns:
        None:
    """
    # for more on formatting Writer documents see,
    # https://python-ooo-dev-tools.readthedocs.io/en/latest/help/writer/format/index.html
    try:
        doc = CalcDoc.from_current_doc()
        create_hyperlinks(doc)
    except Exception as e:
        _ = MsgBox.msgbox(f"This method requires a Calc document.\n{e}")


g_exportedScripts = (make_hyperlinks,)

