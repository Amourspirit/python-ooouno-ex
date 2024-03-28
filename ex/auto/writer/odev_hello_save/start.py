from pathlib import Path
from ooodev.loader import Lo
from ooodev.write import WriteDoc, ZoomKind


def main() -> int:
    # see: https://python-ooo-dev-tools.readthedocs.io/en/latest/src/loader/lo.html#ooodev.loader.Lo.Loader
    with Lo.Loader(Lo.ConnectSocket()) as loader:
        doc = WriteDoc.create_doc(loader=loader, visible=True)

        # These delays are unnecessary. They are here merely to better
        # illustrate this example in action.
        Lo.delay(300)  # small delay before dispatching zoom command
        doc.zoom(ZoomKind.PAGE_WIDTH)

        cursor = doc.get_cursor()
        cursor.goto_end()  # make sure at end of doc before appending
        cursor.append_para(text="Hello LibreOffice.\n")
        # (Underscores are allowed in numeric literals since Python 3.6)
        Lo.delay(1_000)

        cursor.append_para(text="How are you?")
        Lo.delay(2_000)

        tmp = Path.cwd() / "tmp"
        tmp.mkdir(exist_ok=True, parents=True)
        doc.save_doc(fnm=tmp / "hello.odt")
        doc.close_doc()

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
