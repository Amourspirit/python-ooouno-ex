from __future__ import annotations
from pathlib import Path

import uno

from ooodev.dialog.msgbox import (
    MsgBox,
    MessageBoxType,
    MessageBoxButtonsEnum,
    MessageBoxResultsEnum,
)
from ooodev.draw import Draw, ImpressDoc, DrawText
from ooodev.utils.file_io import FileIO
from ooodev.utils.info import Info
from ooodev.utils.lo import Lo
from ooodev.utils.type_var import PathOrStr


class PointsBuilder:
    def __init__(self, points_fnm: PathOrStr) -> None:
        _ = FileIO.is_exist_file(fnm=points_fnm, raise_err=True)
        self._points_fnm = FileIO.get_absolute_path(points_fnm)

    def main(self) -> None:
        loader = Lo.load_office(Lo.ConnectPipe())

        # create Impress page or Draw slide
        try:
            self._report_templates()
            template_name = "Inspiration.otp"  # "Piano.otp"
            template_fnm = Path(Draw.get_slide_template_path(), template_name)
            _ = FileIO.is_exist_file(template_fnm, True)
            doc = ImpressDoc.create_doc_from_template(
                template_path=template_fnm, loader=loader
            )

            self._read_points(doc)

            print(f"Total no. of slides: {len(doc.slides)}")

            doc.set_visible()
            Lo.delay(2000)

            msg_result = MsgBox.msgbox(
                "Do you wish to save document?",
                "Save",
                boxtype=MessageBoxType.QUERYBOX,
                buttons=MessageBoxButtonsEnum.BUTTONS_YES_NO,
            )
            if msg_result == MessageBoxResultsEnum.YES:
                tmp = Path.cwd() / "tmp"
                tmp.mkdir(exist_ok=True)
                doc.save_doc(tmp / "points.odp")
                print(f"Saved to {tmp / 'points.odp'}")

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

    def _report_templates(self) -> None:
        template_dirs = Info.get_dirs(setting="Template")
        print("Templates dir:")
        for dir in template_dirs:
            print(f"  {dir}")

        template_dir = Draw.get_slide_template_path()
        print()
        print(f'Templates files in "{template_dir}"')
        template_fnms = FileIO.get_file_paths(template_dir)
        for fnm in template_fnms:
            print(f"  {fnm}")

    def _read_points(self, doc: ImpressDoc) -> None:
        # Read in a text file of points which are converted to slides.
        # Formatting rules:
        # * ">", ">>", etc are points and their levels
        # * any other lines are the title text of a new slide
        curr_slide = doc.get_slide(idx=0)
        curr_slide.title_slide(
            title="Python-Generated Slides",
            sub_title="Using LibreOffice",
        )
        try:

            def process_bullet(
                line: str, draw_text: DrawText[ImpressDoc] | None
            ) -> None:
                # count the number of '>'s to determine the bullet level
                if draw_text is None:
                    print(f"No slide body for {line}")
                    return

                pos = 0
                s_lst = [*line]
                ch = s_lst[pos]
                while ch == ">":
                    pos += 1
                    ch = s_lst[pos]
                sub_str = "".join(s_lst[pos:]).strip()

                draw_text.add_bullet(level=pos - 1, text=sub_str)

            body: DrawText[ImpressDoc] | None = None
            with open(self._points_fnm, "r") as file:
                # remove empty lines
                data = (row for row in file if row.strip())
                # chain generator
                # strip of remove anything starting //
                # // for comment
                data = (row for row in data if not row.lstrip().startswith("//"))

                for row in data:
                    ch = row[:1]
                    if ch == ">":
                        process_bullet(line=row, draw_text=body)
                    else:
                        curr_slide = doc.add_slide()
                        body = curr_slide.bullets_slide(title=row.strip())
            print(f"Read in point file: {self._points_fnm.name}")
        except Exception as e:
            print(f"Error reading points file: {self._points_fnm}")
            print(f"  {e}")
