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
from ooodev.utils.info import Info
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
            doc_component = Lo.open_doc(self._fnm, loader)

            # slideshow start() crashes if the doc is not visible

            if not Info.is_doc_type(obj=doc_component, doc_type=Lo.Service.IMPRESS):
                print("-- Not a slides presentation")
                Lo.close_office()
                return

            doc = ImpressDoc(doc_component)
            doc.set_visible()

            slides = doc.get_slides()
            num_slides = slides.get_count()
            print(f"No. of slides: {num_slides}")

            # add a title-only slide with a graphic at the end
            last_page = ImpressPage(
                owner=doc, component=slides.insert_new_by_index(num_slides)
            )
            last_page.title_only_slide(header="Any Questions?")
            last_page.draw_image(fnm=self._im_fnm)

            # add a title/subtitle slide at the start
            first_page = ImpressPage(owner=doc, component=slides.insert_new_by_index(0))
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
