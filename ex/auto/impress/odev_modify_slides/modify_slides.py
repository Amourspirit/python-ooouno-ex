from __future__ import annotations

import uno
from ooodev.dialog.msgbox import (
    MsgBox,
    MessageBoxType,
    MessageBoxButtonsEnum,
    MessageBoxResultsEnum,
)
from ooodev.draw import ImpressDoc, ImpressPage
from ooodev.utils.file_io import FileIO
from ooodev.utils.lo import Lo
from ooodev.utils.type_var import PathOrStr


class ModifySlides:
    def __init__(self, fnm: PathOrStr, im_fnm: PathOrStr) -> None:
        _ = FileIO.is_exist_file(fnm=fnm, raise_err=True)
        _ = FileIO.is_exist_file(fnm=im_fnm, raise_err=True)
        self._fnm = FileIO.get_absolute_path(fnm)
        self._im_fnm = FileIO.get_absolute_path(im_fnm)

    def main(self) -> None:
        loader = Lo.load_office(Lo.ConnectPipe())

        try:
            doc = ImpressDoc.open_doc(fnm=self._fnm, loader=loader)
            doc.set_visible()

            num_slides = len(doc.slides)
            print(f"No. of slides: {num_slides}")

            # add a title-only slide with a graphic at the end
            last_page = doc.slides.insert_new_by_index(-1)  # -1 means insert at end
            last_page.title_only_slide(header="Any Questions?")
            _ = last_page.draw_image(fnm=self._im_fnm)

            # add a title/subtitle slide at the start
            first_page = doc.slides.insert_new_by_index(0)
            first_page.title_slide(
                title="Interesting Slides", sub_title="Brought to you by OooDev"
            )

            Lo.delay(2000)
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
