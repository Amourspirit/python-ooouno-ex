#!/usr/bin/env python
from __future__ import annotations
from pathlib import Path
import uno

from ooodev.loader import Lo
from ooodev.calc import CalcDoc, CalcSheet, ZoomKind
from ooodev.utils.color import CommonColor
from ooodev.utils.props import Props


def do_cell_range(sheet: CalcSheet) -> None:
    vals = (
        ("Name", "Fruit", "Quantity"),
        ("Alice", "Apples", 3),
        ("Alice", "Oranges", 7),
        ("Bob", "Apples", 3),
        ("Alice", "Apples", 9),
        ("Bob", "Apples", 5),
        ("Bob", "Oranges", 6),
        ("Alice", "Oranges", 3),
        ("Alice", "Apples", 8),
        ("Alice", "Oranges", 1),
        ("Bob", "Oranges", 2),
        ("Bob", "Oranges", 7),
        ("Bob", "Apples", 1),
        ("Alice", "Apples", 8),
        ("Alice", "Oranges", 8),
        ("Alice", "Apples", 7),
        ("Bob", "Apples", 1),
        ("Bob", "Oranges", 9),
        ("Bob", "Oranges", 3),
        ("Alice", "Oranges", 4),
        ("Alice", "Apples", 9),
    )
    with sheet.calc_doc:
        # use doc context manager to lock controllers for faster updates.
        sheet.set_array(values=vals, name="A1:C21")  # or just "A1"
        cell = sheet.get_cell(cell_name="A22")
        cell.set_val("Total")

        cell = sheet.get_cell(cell_name="C22")
        cell.set_val("=SUM(C2:C21)")

        # set Border around data and summary.
        rng = sheet.get_range(range_name="A1:C22")
        rng.style_borders_sides(color=CommonColor.LIGHT_BLUE, width=2.85)


def set_original_prop(sheet: CalcSheet) -> None:
    rng = sheet.rng("B2:C21")
    for cell_obj in rng:
        cell = sheet[cell_obj]
        cell.set_custom_property("OriginalValue", cell.value)


def get_docs_dir() -> Path:
    return Path(__file__).parent / "data" / "src_doc"


def main() -> int:
    _ = Lo.load_office(Lo.ConnectSocket())
    doc = None
    try:
        doc = CalcDoc.create_doc(visible=True)
        events = doc.component.getEvents()
        if events.hasByName("OnViewCreated"):
            element_props = events.getByName("OnViewCreated")  # Tuple of Property Value
            if element_props:
                dprops = Props.data_to_dict(element_props)
            else:
                dprops = {}
            dprops["Script"] = (
                "vnd.sun.star.script:custom_props.py$register_custom_prop_interceptor?language=Python&location=document"
            )
            dprops["EventType"] = "Script"
            data = Props.make_props(**dprops)
            uno_data = uno.Any("[]com.sun.star.beans.PropertyValue", data)
            uno.invoke(events, "replaceByName", ("OnViewCreated", uno_data))
        clean_events = ("OnCopyTo", "OnSaveAs", "OnSave")
        for clean_event in clean_events:
            if events.hasByName(clean_event):
                element_props = events.getByName(clean_event)  # Tuple of Property Value
                dprops = {}
                dprops["Script"] = (
                    "vnd.sun.star.script:custom_props.py$clean_custom_props?language=Python&location=document"
                )
                dprops["EventType"] = "Script"
                data = Props.make_props(**dprops)
                uno_data = uno.Any("[]com.sun.star.beans.PropertyValue", data)
                uno.invoke(events, "replaceByName", (clean_event, uno_data))

        sheet = doc.sheets[0]
        # doc.set_visible()
        # wait for document to be ready
        Lo.delay(300)
        doc.zoom(ZoomKind.ENTIRE_PAGE)

        do_cell_range(sheet=sheet)
        doc.freeze_rows(num_rows=1)
        set_original_prop(sheet=sheet)

        doc_path = get_docs_dir() / "src_doc.ods"
        if doc_path.exists():
            doc_path.unlink()
        try:
            doc.save_doc(doc_path)
        except Exception as e:
            print(e)

    except Exception as e:
        print(e)
        raise
    finally:
        if doc:
            doc.close_doc()
        Lo.close_office()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
