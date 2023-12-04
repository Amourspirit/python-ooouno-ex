#
# on wayland (some versions of Linux)
# may get error:
#    (soffice:67106): Gdk-WARNING **: 02:35:12.168: XSetErrorHandler() called with a GDK error trap pushed. Don't do that.
# This seems to be a Wayland/Java compatibility issues.
# see: http://www.babelsoft.net/forum/viewtopic.php?t=24545
from __future__ import annotations
from pathlib import Path
import uno
from ooodev.utils.lo import Lo
from ooodev.calc import Calc
from ooodev.calc import CalcDoc
from ooodev.utils.gui import GUI
from tab_dialog import Tabs


def main() -> int:
    fnm = Path(__file__).parent / "data" / "sales_data.ods"
    with Lo.Loader(Lo.ConnectSocket(), opt=Lo.Options(verbose=True)):
        doc = CalcDoc(Calc.open_doc(fnm=fnm))

        GUI.set_visible(visible=True, doc=doc.component)
        tabs = Tabs(doc=doc)
        Lo.delay(500)
        tabs.show()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
