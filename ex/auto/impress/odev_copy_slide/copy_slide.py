from __future__ import annotations

import uno
from ooodev.dialog.msgbox import (
    MsgBox,
    MessageBoxType,
    MessageBoxButtonsEnum,
    MessageBoxResultsEnum,
)
from ooodev.utils.dispatch.draw_view_dispatch import DrawViewDispatch
from ooodev.utils.dispatch.draw_drawing_dispatch import DrawDrawingDispatch
from ooodev.utils.dispatch.global_edit_dispatch import GlobalEditDispatch
from ooodev.draw import ImpressDoc
from ooodev.utils.file_io import FileIO
from ooodev.utils.lo import Lo
from ooodev.utils.type_var import PathOrStr


class CopySlide:
    def __init__(self, fnm: PathOrStr, from_idx: int, to_idx: int) -> None:
        idx1 = int(from_idx)
        idx2 = int(to_idx)
        if idx1 < 0:
            raise IndexError("From Index must be greater or equal to 0")
        if idx2 < 0:
            raise IndexError("To Index must be greater or equal to 0")
        self._from_idx = idx1
        self._to_idx = idx2
        FileIO.is_exist_file(fnm=fnm, raise_err=True)
        self._fnm = FileIO.get_absolute_path(fnm)

    def main(self) -> None:
        loader = Lo.load_office(Lo.ConnectPipe())

        try:
            doc = ImpressDoc(Lo.open_doc(fnm=self._fnm, loader=loader))
            num_slides = doc.get_slides_count()
            if self._from_idx >= num_slides or self._to_idx >= num_slides:
                Lo.close_office()
                raise IndexError("One or both indices are out of range")

            doc.set_visible()

            self._copy_to(doc=doc)

            # Draw.delete_slide(doc=doc, idx=self._from_idx)
            # a problem if the copying changes the indices

            # Lo.save(doc) # overwrites original

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
        except IndexError:
            raise
        except Exception:
            Lo.close_office()
            raise

    def _copy_to(self, doc: ImpressDoc) -> None:
        # Copy fromIdx slide to the clipboard in slide-sorter mode,
        # then paste it to after the toIdx slide.

        # Switch to slide sorter view so that slides can be pasted
        Lo.delay(1000)
        Lo.dispatch_cmd(cmd=DrawViewDispatch.DIA_MODE)

        # give Office a few seconds of time to do it
        Lo.delay(3000)

        from_slide = doc.get_slide(idx=self._from_idx)
        to_slide = doc.get_slide(idx=self._to_idx)
        doc.goto_page(from_slide.component)
        Lo.dispatch_cmd(cmd=GlobalEditDispatch.COPY)
        Lo.delay(500)
        print(f"Copied {self._from_idx}")

        doc.goto_page(to_slide.component)
        Lo.delay(500)
        Lo.dispatch_cmd(GlobalEditDispatch.PASTE)
        Lo.delay(500)
        print(f"Paste to after {self._to_idx}")

        # Lo.dispatchCmd("PageMode");  // back to normal mode (not working)
        Lo.dispatch_cmd(cmd=DrawDrawingDispatch.DRAWING_MODE)
        Lo.delay(500)
