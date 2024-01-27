from __future__ import annotations
from pathlib import Path
from ooodev.utils.lo import Lo
from ooodev.write import WriteDoc, ZoomKind


def main() -> int:
    # see: https://python-ooo-dev-tools.readthedocs.io/en/latest/src/utils/lo.html#ooodev.utils.lo.Lo.Loader
    with Lo.Loader(Lo.ConnectSocket()) as loader:
        doc = WriteDoc.create_doc(loader)

        doc.set_visible()
        Lo.delay(300)  # small delay before dispatching zoom command
        doc.zoom(ZoomKind.PAGE_WIDTH)

        cursor = doc.get_cursor()
        cursor.goto_end()  # make sure at end of doc before appending
        cursor.append_para(text="Hello LibreOffice.\n")
        Lo.delay(1_000)  # Slow things down so user can see

        cursor.append_para(text="How are you?")
        Lo.delay(2_000)
        tmp = Path.cwd() / "tmp"
        tmp.mkdir(exist_ok=True, parents=True)
        doc.save_doc(fnm=tmp / "hello.odt")
        doc.close_doc()

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
