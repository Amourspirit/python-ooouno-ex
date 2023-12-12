from __future__ import annotations

import uno
from ooodev.draw import Draw, ImpressDoc
from ooodev.utils.dispatch.draw_view_dispatch import DrawViewDispatch
from ooodev.utils.file_io import FileIO
from ooodev.utils.lo import Lo
from ooodev.utils.props import Props
from ooodev.utils.type_var import PathOrStr


class BasicShow:
    def __init__(self, fnm: PathOrStr) -> None:
        _ = FileIO.is_exist_file(fnm=fnm, raise_err=True)
        self._fnm = FileIO.get_absolute_path(fnm)

    def main(self) -> None:
        with Lo.Loader(Lo.ConnectPipe()) as loader:
            doc = ImpressDoc(Lo.open_doc(fnm=self._fnm, loader=loader))
            try:
                # slideshow start() crashes if the doc is not visible
                doc.set_visible()

                show = doc.get_show()
                Props.show_obj_props("Slide show", show)

                Lo.delay(500)
                Lo.dispatch_cmd(DrawViewDispatch.PRESENTATION)
                # show.start() starts slideshow but not necessarily in 100% full screen
                # show.start()

                sc = doc.get_show_controller()
                Draw.wait_ended(sc)

            finally:
                doc.close_doc()
