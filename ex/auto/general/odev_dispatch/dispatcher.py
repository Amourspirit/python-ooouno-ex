from __future__ import annotations

import uno
from ooodev.dialog.msgbox import (
    MsgBox,
    MessageBoxType,
    MessageBoxButtonsEnum,
    MessageBoxResultsEnum,
)
from ooodev.exceptions.ex import CancelEventError
from ooodev.loader import Lo
from ooodev.office.write import Write
from ooodev.utils.file_io import FileIO
from ooodev.utils.gui import GUI
from ooodev.utils.type_var import PathOrStr

from ooo.dyn.presentation.animation_speed import AnimationSpeed as AnimationSpeed
from ooo.dyn.presentation.fade_effect import FadeEffect as FadeEffect

try:
    # only windows
    from odevgui_win.robot_keys import RobotKeys, SendKeyInfo
    from odevgui_win.keys.writer_key_codes import WriterKeyCodes
except ImportError:
    RobotKeys, SendKeyInfo, WriterKeyCodes = None, None, None


class Dispatcher:
    def __init__(self, fnm: PathOrStr) -> None:
        _ = FileIO.is_exist_file(fnm=fnm, raise_err=True)
        self._fnm = fnm

    def main(self) -> None:
        loader = Lo.load_office(Lo.ConnectPipe())

        try:
            doc = Write.open_doc(self._fnm, loader)

            # slideshow start() crashes if the doc is not visible
            GUI.set_visible(visible=True, doc=doc)
            Lo.delay(1_000)  # give time of zoom may no work.
            GUI.zoom(GUI.ZoomEnum.ZOOM_75_PERCENT)
            Lo.delay(1_000)

            # put doc into readonly mode
            Lo.dispatch_cmd("ReadOnlyDoc")
            Lo.delay(500)

            # self._toggle_side_bar()

            # opens get involved webpage of LibreOffice in local browser
            Lo.dispatch_cmd("GetInvolved")
            try:
                Lo.dispatch_cmd("About")
            except CancelEventError as e:
                print(e)

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

    def _toggle_side_bar(self) -> None:
        # RobotKeys is currently windows only
        if not RobotKeys:
            Lo.print("odevgui_win not found.")
            return
        # see: https://ooo-dev-tools-gui-win.readthedocs.io/en/latest/src/robot_keys.html
        # # send a predefined key sequence
        RobotKeys.send_current(SendKeyInfo(WriterKeyCodes.KB_SIDE_BAR))
