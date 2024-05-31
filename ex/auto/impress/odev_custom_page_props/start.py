from __future__ import annotations
from ooodev.loader import Lo
from ooodev.draw import ImpressDoc, ZoomKind
from ooodev.utils.helper.dot_dict import DotDict


def main() -> int:
    with Lo.Loader(Lo.ConnectSocket()) as loader:
        doc = ImpressDoc.create_doc(loader=loader, visible=True)

        Lo.delay(300)  # small delay before dispatching zoom command

        slide = doc.slides[0]
        slide.draw_rectangle(10, 10, 14, 12)

        slide.set_custom_property("Prop1", "Hello")
        slide.set_custom_property("Prop2", "World")
        slide.set_custom_property("Prop3", 777)

        dd = DotDict()
        dd.Meta1 = "Some meta data"
        dd.Meta2 = "Some other meta data"
        slide.set_custom_properties(dd)

        cprops = slide.get_custom_properties()
        for key, value in cprops.items():
            print(f"{key}: {value}")

        tmp_file = Lo.tmp_dir / "props.odp"
        print("Saving document with custom properties")
        print(f"File: {tmp_file}")
        doc.save_doc(fnm=tmp_file)
        doc.close_doc()

        doc = None
        doc = ImpressDoc.open_doc(fnm=tmp_file, loader=loader, visible=True)
        print()
        print("Opened document again. Custom properties:")
        slide = doc.slides[0]
        cprops = slide.get_custom_properties()
        for key, value in cprops.items():
            print(f"{key}: {value}")
        doc.close()
        doc = None

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
