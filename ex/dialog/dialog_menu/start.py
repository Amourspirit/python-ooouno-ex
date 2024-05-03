from __future__ import annotations
import uno
from ooodev.loader import Lo
from ooodev.calc import CalcDoc
import macro_code


def main() -> int:
    with Lo.Loader(Lo.ConnectSocket(), opt=Lo.Options(verbose=True)):
        doc = CalcDoc.create_doc(visible=True)

        Lo.delay(300)
        doc.zoom_value(100)
        macro_code.show_dialog()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
