from __future__ import annotations
from ooodev.loader import Lo
from ooodev.write import WriteDoc, ZoomKind
from ooodev.utils.helper.dot_dict import DotDict


def main() -> int:
    # see: https://python-ooo-dev-tools.readthedocs.io/en/latest/src/loader/lo.html#ooodev.loader.Lo.Loader
    with Lo.Loader(Lo.ConnectSocket()) as loader:
        doc = WriteDoc.create_doc(loader=loader, visible=True)

        # These delays are unnecessary. They are here merely to better
        # illustrate this example in action.
        Lo.delay(300)  # small delay before dispatching zoom command
        doc.zoom(ZoomKind.PAGE_WIDTH)

        cursor = doc.get_cursor()
        cursor.append_para("Hello LibreOffice")
        doc.set_custom_property("Prop1", "Hello")
        doc.set_custom_property("Prop2", "World")
        doc.set_custom_property("Prop3", 777)

        dd = DotDict()
        dd.Meta1 = "Some meta data"
        dd.Meta2 = "Some other meta data"
        doc.set_custom_properties(dd)

        cprops = doc.get_custom_properties()
        for key, value in cprops.items():
            print(f"{key}: {value}")

        tmp_file = Lo.tmp_dir / "props.odt"
        print("Saving document with custom properties")
        print(f"File: {tmp_file}")
        doc.save_doc(fnm=tmp_file)
        doc.close_doc()

        doc = None
        doc = WriteDoc.open_doc(fnm=tmp_file, loader=loader, visible=True)
        print()
        print("Opened document again. Custom properties:")
        cprops = doc.get_custom_properties()
        for key, value in cprops.items():
            print(f"{key}: {value}")
        doc.close()
        doc = None

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
