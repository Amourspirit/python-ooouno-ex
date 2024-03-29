from __future__ import annotations
from typing import TYPE_CHECKING

import uno
from com.sun.star.drawing import XDrawPages
from com.sun.star.lang import XComponent

from ooodev.dialog.msgbox import (
    MsgBox,
    MessageBoxType,
    MessageBoxButtonsEnum,
    MessageBoxResultsEnum,
)
from ooodev.office.draw import Draw, mEx
from ooodev.loader import Lo
from ooodev.utils.dispatch.draw_drawing_dispatch import DrawDrawingDispatch
from ooodev.utils.dispatch.draw_view_dispatch import DrawViewDispatch
from ooodev.utils.dispatch.global_edit_dispatch import GlobalEditDispatch
from ooodev.utils.file_io import FileIO
from ooodev.utils.gui import GUI
from ooodev.utils.type_var import PathOrStr


try:
    # only in windows
    from odevgui_win.dialog_auto import DialogAuto
except ImportError:
    DialogAuto = None

if TYPE_CHECKING:
    # these import are only being used for typing
    # therefore not needed at runtime.
    # from __future__ import annotations takes care of the rest.
    from com.sun.star.drawing import XDrawPage
    from com.sun.star.frame import XController
    from com.sun.star.frame import XFrame


class AppendSlides:
    def __init__(self, *fnms: PathOrStr) -> None:
        if len(fnms) == 0:
            raise ValueError("At lease one file is required. fnms has no values.")
        for fnm in fnms:
            _ = FileIO.is_exist_file(fnm, True)

        self._fnms = fnms

    def main(self) -> None:
        loader = Lo.load_office(Lo.ConnectPipe())

        try:
            doc = Lo.open_doc(fnm=self._fnms[0], loader=loader)

            GUI.set_visible(visible=True, doc=doc)

            self._to_ctrl = GUI.get_current_controller(doc)
            self._to_frame = GUI.get_frame(doc)

            # Switch to slide sorter view so that slides can be pasted
            Lo.delay(500)
            Lo.dispatch_cmd(cmd=DrawViewDispatch.DIA_MODE, frame=self._to_frame)

            to_slides = Draw.get_slides(doc)

            # monitor for Confirmation dialog
            if DialogAuto:
                DialogAuto.monitor_dialog("y")

            for fnm in self._fnms[1:]:  # start at 1
                try:
                    app_doc = Lo.open_doc(fnm=fnm, loader=loader)
                except Exception as e:
                    print(f'Could not open the file "{fnm}"')
                    print(f"  {e}")
                    continue

                self._append_doc(to_slides=to_slides, doc=app_doc)

            Lo.delay(500)
            # Lo.dispatch_cmd(cmd=DrawViewDispatch.PAGE_MODE, frame=self._to_frame)  # does not work
            Lo.dispatch_cmd(cmd=DrawDrawingDispatch.DRAWING_MODE)
            Lo.delay(1000)

            msg_result = MsgBox.msgbox(
                "Do you wish to close document?",
                "All done",
                boxtype=MessageBoxType.QUERYBOX,
                buttons=MessageBoxButtonsEnum.BUTTONS_YES_NO,
            )
            if msg_result == MessageBoxResultsEnum.YES:
                Lo.close_doc(doc=doc, deliver_ownership=True)
                Lo.close_office()
            else:
                print("Keeping document open")
        except Exception:
            Lo.close_office()
            raise

    def _append_doc(self, to_slides: XDrawPages, doc: XComponent) -> None:
        # Append doc to the end of  toSlides.
        # Access the slides in the document, and the document's controller and frame refs.
        # Switch to slide sorter view so that slides can be copied.
        GUI.set_visible(visible=True, doc=doc)

        from_ctrl = GUI.get_current_controller(doc)
        from_frame = GUI.get_frame(doc)
        Lo.dispatch_cmd(cmd="DiaMode", frame=from_frame)
        try:
            from_slides = Draw.get_slides(doc)
            print("- Adding slides")
            self._append_slides(
                to_slides=to_slides,
                from_slides=from_slides,
                from_ctrl=from_ctrl,
                from_frame=from_frame,
            )
        except mEx.DrawPageMissingError:
            print("- No Slides Found")

        # Lo.dispatchCmd("PageMode");  // back to normal mode (not working)
        Lo.dispatch_cmd(cmd=DrawDrawingDispatch.DRAWING_MODE)
        Lo.close_doc(doc)
        print()

    def _append_slides(
        self,
        to_slides: XDrawPages,
        from_slides: XDrawPages,
        from_ctrl: XController,
        from_frame: XFrame,
    ) -> None:
        # Append fromSlides to the end of toSlides
        # Loop through the fromSlides, copying each one.
        for i in range(from_slides.getCount()):
            from_slide = Draw.get_slide(from_slides, i)

            # the copy will be placed after this slide
            to_slide = Draw.get_slide(to_slides, to_slides.getCount() - 1)

            self._copy_to(
                from_slide=from_slide,
                from_ctrl=from_ctrl,
                from_frame=from_frame,
                to_slide=to_slide,
                to_ctrl=self._to_ctrl,
                to_frame=self._to_frame,
            )

    def _copy_to(
        self,
        from_slide: XDrawPage,
        from_ctrl: XController,
        from_frame: XFrame,
        to_slide: XDrawPage,
        to_ctrl: XController,
        to_frame: XFrame,
    ) -> None:
        # Copy fromSlide to the clipboard, and
        # then paste it to after the toSlide. Unfortunately, the
        # paste requires a "Yes" button to be pressed.

        Draw.goto_page(from_ctrl, from_slide)  # select this slide
        print("-- Copy -->")
        Lo.dispatch_cmd(cmd=GlobalEditDispatch.COPY, frame=from_frame)
        Lo.delay(1000)

        Draw.goto_page(to_ctrl, to_slide)
        print("Paste")

        # needs automation at this point to monitor for dialog and click the dialog button.
        # for windows use ooo-dev-tools-gui-win - https://ooo-dev-tools-gui-win.readthedocs.io/en/latest/index.html

        Lo.dispatch_cmd(cmd=GlobalEditDispatch.PASTE, frame=to_frame)
