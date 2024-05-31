from __future__ import annotations
from ooodev.loader import Lo
from ooodev.calc import CalcDoc, ZoomKind
from ooodev.utils.helper.dot_dict import DotDict


def main() -> int:
    with Lo.Loader(Lo.ConnectSocket()) as loader:
        doc = CalcDoc.create_doc(loader=loader, visible=True)

        Lo.delay(300)  # small delay before dispatching zoom command
        doc.zoom(ZoomKind.PAGE_WIDTH)

        sheet = doc.sheets[0]

        cell = sheet["A1"]
        cell.value = "Hello LibreOffice"
        sheet.set_custom_property("Prop1", "Hello")
        sheet.set_custom_property("Prop2", "World")
        sheet.set_custom_property("Prop3", 777)

        dd = DotDict()
        dd.Meta1 = "Some meta data"
        dd.Meta2 = "Some other meta data"
        sheet.set_custom_properties(dd)

        sheet2 = doc.sheets.insert_sheet("Sheet2")
        sheet_props = DotDict(
            SheetPurpose="Save a planet", SheetOwner="Elon Musk", GoalYear=2030
        )
        sheet2.set_custom_properties(sheet_props)

        print("Custom properties for Sheet1:")
        sheet1_props = sheet.get_custom_properties()
        for key, value in sheet1_props.items():
            print(f"{key}: {value}")

        print("Custom properties for Sheet2:")
        sheet2_props = sheet2.get_custom_properties()
        for key, value in sheet2_props.items():
            print(f"{key}: {value}")

        tmp_file = Lo.tmp_dir / "props.ods"
        print("Saving document with custom properties")
        print(f"File: {tmp_file}")
        doc.save_doc(fnm=tmp_file)
        doc.close_doc()

        doc = None
        doc = CalcDoc.open_doc(fnm=tmp_file, loader=loader, visible=True)
        print()
        print("Opened document again. Custom properties:")
        print("Custom properties for Sheet2:")
        sheet1 = doc.sheets[0]
        sheet1_props = sheet1.get_custom_properties()
        for key, value in sheet2_props.items():
            print(f"{key}: {value}")

        print("Custom properties for Sheet2:")
        sheet2 = doc.sheets[1]
        sheet2_props = sheet2.get_custom_properties()
        for key, value in sheet2_props.items():
            print(f"{key}: {value}")
        doc.close()
        doc = None

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
