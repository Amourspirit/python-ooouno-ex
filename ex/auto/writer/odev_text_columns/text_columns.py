from __future__ import annotations
import uno
from com.sun.star.container import XNamed
from com.sun.star.text import XTextColumns
from com.sun.star.text import XTextContent

from ooodev.utils.lo import Lo
from ooodev.utils.props import Props
from ooodev.write import Write, WriteDoc

from ooodev.dialog.msgbox import (
    MsgBox,
    MessageBoxType,
    MessageBoxButtonsEnum,
    MessageBoxResultsEnum,
)

# Python Example of https://wiki.documentfoundation.org/Documentation/DevGuide/Text_Documents#Columns


class TextColumns:
    def __init__(self) -> None:
        pass

    def main(self) -> None:
        loader = Lo.load_office(Lo.ConnectSocket())
        try:
            doc = WriteDoc(Write.create_doc(loader=loader))
            doc.set_visible()

            cursor = doc.get_cursor()
            # insert a new paragraph
            cursor.append_para()

            # insert the string 'I am a fish.' 100 times
            sb = []
            for _ in range(100):
                sb.append("I am a fish.")

            cursor.append(" ".join(sb))

            # insert a paragraph break after the text
            cursor.append_para()

            # Cursor is also XParagraphCursor interface of our text
            # Jump back before all the text we just inserted
            cursor.goto_previous_paragraph()
            cursor.goto_previous_paragraph()

            # Insert a string at the beginning of the block of text
            # not using cursor.append here because append moves the cursor to the end of the document text.
            # cursor.append("Fish section begins:")
            cursor.set_string("Fish section begins:")  # does not move the cursor

            # Then select the I am fish content of the text we just inserted
            cursor.goto_next_paragraph()
            cursor.goto_next_paragraph(True)

            # Create a new text section and get it's XNamed interface
            named_section = Lo.create_instance_msf(
                XNamed, "com.sun.star.text.TextSection", raise_err=True
            )

            # Set the name of our new section (appropriately) to 'Fish'
            named_section.setName("Fish")

            # Create the TextColumns service and get it's XTextColumns interface
            columns = Lo.create_instance_msf(
                XTextColumns, "com.sun.star.text.TextColumns", raise_err=True
            )

            # We want three columns
            columns.setColumnCount(3)

            # Get the TextColumns, and make the middle one narrow with a larger margin
            # on the left than the right
            current_columns = columns.getColumns()
            column = current_columns[1]
            column.Width = round(column.Width / 2)
            column.LeftMargin = 350  # 1/100th mm
            column.RightMargin = 200  # 1/100th mm
            # Set the updated TextColumns back to the XTextColumns
            columns.setColumns(current_columns)

            # Set the columns to the Text Section
            Props.set(named_section, TextColumns=columns)

            # Insert the 'Fish' section over the currently selected text
            write_text = cursor.get_write_text()
            write_text.insert_text_content(content=named_section, absorb=True)

            # Create a new empty paragraph and get it's XTextContent interface
            para = Lo.create_instance_msf(
                XTextContent, "com.sun.star.text.Paragraph", raise_err=True
            )

            # Insert the empty paragraph after the fish Text Section
            # write_text contains methods for XRelativeTextContentInsert interface.
            write_text.insert_text_content_after(
                new_content=para, predecessor=named_section
            )

            msg_result = MsgBox.msgbox(
                "Do you wish to close document?",
                "All done",
                boxtype=MessageBoxType.QUERYBOX,
                buttons=MessageBoxButtonsEnum.BUTTONS_YES_NO,
            )
            if msg_result == MessageBoxResultsEnum.YES:
                doc.close_doc()
                Lo.close_office()
            else:
                print("Keeping document open")

        except Exception:
            Lo.close_office()
            raise
