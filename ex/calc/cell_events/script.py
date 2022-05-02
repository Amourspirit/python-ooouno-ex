# coding: utf-8
"""
Demonstrates Getting info on cell selection and cell changed events.


event supportedInterfaces={
    com.sun.star.beans.XPropertySet,
    com.sun.star.beans.XMultiPropertySet,
    com.sun.star.beans.XPropertyState,
    com.sun.star.sheet.XSheetOperation,
    com.sun.star.chart.XChartDataArray,
    com.sun.star.util.XIndent,
    com.sun.star.sheet.XCellRangesQuery,
    com.sun.star.sheet.XFormulaQuery,
    com.sun.star.util.XReplaceable,
    com.sun.star.util.XModifyBroadcaster,
    com.sun.star.lang.XServiceInfo,
    com.sun.star.lang.XUnoTunnel,
    com.sun.star.beans.XTolerantMultiPropertySet,
    com.sun.star.lang.XTypeProvider,
    com.sun.star.uno.XWeak,
    com.sun.star.sheet.XCellRangeAddressable,
    com.sun.star.sheet.XSheetCellRange,
    com.sun.star.sheet.XArrayFormulaRange,
    com.sun.star.sheet.XArrayFormulaTokens,
    com.sun.star.sheet.XCellRangeData,
    com.sun.star.sheet.XCellRangeFormula,
    com.sun.star.sheet.XMultipleOperation,
    com.sun.star.util.XMergeable,
    com.sun.star.sheet.XCellSeries,
    com.sun.star.table.XAutoFormattable,
    com.sun.star.util.XSortable,
    com.sun.star.sheet.XSheetFilterableEx,
    com.sun.star.sheet.XSubTotalCalculatable,
    com.sun.star.table.XColumnRowRange,
    com.sun.star.util.XImportable,
    com.sun.star.sheet.XCellFormatRangesSupplier,
    com.sun.star.sheet.XUniqueCellFormatRangesSupplier,
    com.sun.star.table.XCell,
    com.sun.star.sheet.XCellAddressable,
    com.sun.star.text.XText,
    com.sun.star.container.XEnumerationAccess,
    com.sun.star.sheet.XSheetAnnotationAnchor,
    com.sun.star.text.XTextFieldsSupplier,
    com.sun.star.document.XActionLockable,
    com.sun.star.sheet.XFormulaTokens,
    com.sun.star.table.XCell2
    }
"""
from __future__ import annotations
from typing import TYPE_CHECKING
import json
import scriptforge as SF
if TYPE_CHECKING:
    from com.sun.star.uno import XInterface
    from com.sun.star.form.component import RichTextControl
    from com.sun.star.sheet import XSpreadsheet


def OnSheetContentChange(e: XInterface) -> None:
    doc: SF.SFDocuments.SF_Calc = SF.CreateScriptService("Calc")
    sht: XSpreadsheet = e.Spreadsheet
    
    txt:RichTextControl = sht.DrawPage.Forms[0]["txtChanged"]
    val = doc.GetValue(e.AbsoluteName)
    if isinstance(val, int) and val < 0:
        e.setValue(0.0)
    d = _get_cell_info(e, doc)    
    txt.setString(json.dumps(d, sort_keys=True, indent=4))

def OnSelectionChanged(e: XInterface) -> None:
    doc: SF.SFDocuments.SF_Calc = SF.CreateScriptService("Calc")
    sht: XSpreadsheet = e.Spreadsheet
    
    txt:RichTextControl = sht.DrawPage.Forms[0]["txtChangedSel"]
    val = doc.GetValue(e.AbsoluteName)
    if isinstance(val, int) and val < 0:
        e.setValue(0.0)
    d = _get_cell_info(e, doc)    
    txt.setString(json.dumps(d, sort_keys=True, indent=4))

def _get_cell_info(e: XInterface, doc: SF.SFDocuments.SF_Calc) -> dict:
    cell = doc.FirstCell(e.AbsoluteName)
    val = doc.GetValue(cell)
    d = {
        "AbsoluteName": e.AbsoluteName,
        "Value": val,
        "FirstColumn": doc.FirstColumn(e.AbsoluteName),
        "FirstRow": doc.FirstRow(e.AbsoluteName),
        "LastRow": doc.LastColumn(e.AbsoluteName),
        "FirstRow": doc.LastRow(e.AbsoluteName),
        "Type": type(val).__name__,
        "IsMerged": e.IsMerged,
        "IsTextWrapped": e.IsTextWrapped,
        "IsCellBackgroundTransparent": e.IsCellBackgroundTransparent,
        "NotANumber": not e.NotANumber,
        "NumberFormat": e.NumberFormat,
        "RangeAddress": {
            "StartColumn": e.RangeAddress.StartColumn,
            "EndColumn": e.RangeAddress.EndColumn,
            "StartRow": e.RangeAddress.StartRow,
            "EndRow": e.RangeAddress.EndRow
        }
    }
    return d

def console(*args, **kwargs) -> None:
    serv: SF.SFScriptForge.SF_Exception = SF.CreateScriptService('ScriptForge.Exception')
    serv.PythonShell({**globals(), **locals()})

g_exportedScripts = (OnSheetContentChange, OnSelectionChanged, console)
    