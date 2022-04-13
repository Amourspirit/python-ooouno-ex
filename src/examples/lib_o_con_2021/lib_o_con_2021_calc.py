'''
Module to demonstrate SourceForge Methods in the Calc Service	
	
Activate	in create_sheet_example
ClearAll	in clear_contents_v1
ClearFormats	in clear_contents_v2
ClearValues	in clear_contents_v3
CopySheet	in copy_sheet_example
CopySheetFromFile	in copy_from_file_example
CopyToCell	in copy_cells_v1
CopyToRange	in copy_cells_v2
DAvg	in calculate_average
DCount	see Davg
DMax	see Davg
DMin	see Davg
DSum	see Davg
Forms	
GetColumnName	see GetValue
GetFormula	see GetValue
GetValue	in mark_invalid
ImportFromCSVFile	in open_csv_file_v1
ImportFromDatabase	see ImportFromCSVFile
InsertSheet	in create_sheet_example
MoveRange	see CopyToRange
MoveSheet	see InsertSheet 
Offset	in create_random_matrix_v1
RemoveSheet	in remove_sheet_example
RenameSheet	see InsertSheet 
SetArray	in create_random_matrix_v2
SetValue	in create_random_matrix_v1
SetCellStyle	in mark_invalid
SetFormula	see SetValue
SortRange	'''
from ooo.lo.sheet.x_spreadsheet import XSpreadsheet
from ooo.lo.sheet.x_sheet_cell_cursor import XSheetCellCursor
from ooo.lo.sheet.sheet_cell_range import SheetCellRange
from ooo.lo.frame.x_model import XModel

from typing import Union
import scriptforge as SF
import random as rnd
from pathlib import Path

# Creates a 6x6 matrix starting at A1
def create_random_matrix_v1(args=None):
    doc: SF.SFDocuments.SF_Calc = SF.CreateScriptService("Calc")
    for i in range(6):
        for j in range(6):
            target_cell = doc.Offset("A1", i, j)
            r = rnd.random()
            if r < 0.5:
                doc.SetValue(target_cell, "EVEN")
            else:
                doc.SetValue(target_cell, "ODD")

# Creates a 6x6 matrix starting at A1
# Uses the method setArray to insert values
def create_random_matrix_v2(args=None):
    doc: SF.SFDocuments.SF_Calc = SF.CreateScriptService("Calc")
    rnd_word = lambda : "EVEN" if rnd.random() < 0.5 else "ODD"
    values = [[rnd_word() for _ in range(6)] for _ in range(6)]
    doc.SetArray("A1", values)

# Creates an mxn matrix starting at A1 and asks the desired size
def create_random_matrix_v3(args=None):
    doc: SF.SFDocuments.SF_Calc = SF.CreateScriptService("Calc")
    bas: SF.SFScriptForge.SF_Basic = SF.CreateScriptService("Basic")
    n_rows = bas.InputBox("Number of rows")
    n_cols = bas.InputBox("Number of columns")
    for i in range(int(n_rows)):
        for j in range(int(n_cols)):
            target_cell = doc.Offset("A1", i, j)
            r = rnd.random()
            if r < 0.5:
                doc.SetValue(target_cell, "EVEN")
            else:
                doc.SetValue(target_cell, "ODD")

# Creates an mxn matrix starting at A1 and asks the desired size
# Uses the method setArray to insert values
def create_random_matrix_v4(args=None):
    doc: SF.SFDocuments.SF_Calc = SF.CreateScriptService("Calc")
    bas: SF.SFScriptForge.SF_Basic = SF.CreateScriptService("Basic")
    n_rows = bas.InputBox("Number of rows")
    n_cols = bas.InputBox("Number of columns")
    rnd_word = lambda : "EVEN" if rnd.random() < 0.5 else "ODD"
    values = [[rnd_word() for _ in range(int(n_cols))]
              for _ in range(int(n_rows))]
    doc.setArray("A1", values)

# Clear region starting at A1
def clear_region_a1(args=None):
    doc: SF.SFDocuments.SF_Calc = SF.CreateScriptService("Calc")
    cur_sheet: XSpreadsheet  = XSCRIPTCONTEXT.getDocument().CurrentController.ActiveSheet
    cell = cur_sheet.getCellRangeByName("A1")
    cursor: Union[SheetCellRange, XSheetCellCursor] = cur_sheet.createCursorByRange(cell)
    cursor.collapseToCurrentRegion()
    doc.ClearAll(cursor.AbsoluteName)

# Creates a matrix of size 10x8 with random integers between -20 and 100
def create_values_for_example_2(args=None):
    clear_region_a1()
    doc: SF.SFDocuments.SF_Calc = SF.CreateScriptService("Calc")
    data = [[rnd.randint(-20, 100) for _ in range(8)] for _ in range(10)]
    doc.SetArray("A1", data)

# Marks cells with negative values as INVALID and apply the 'Bad' cell style
def mark_invalid(args=None):
    doc: SF.SFDocuments.SF_Calc = SF.CreateScriptService("Calc")
    # Gets address of current selection
    cur_selection: str = doc.CurrentSelection
    # Gets address of first cell in the selection
    _ = doc.Offset(cur_selection, 0, 0, 1, 1)
    for i in range(doc.Height(cur_selection)):
        for j in range(doc.Width(cur_selection)):
            cell = doc.Offset(cur_selection, i, j, 1, 1)
            value = doc.GetValue(cell)
            if value < 0:
                doc.SetValue(cell, "INVALID")
                doc.SetCellStyle(cell, "Bad")

# Example of using ClearAll
def clear_contents_v1(args=None):
    doc: SF.SFDocuments.SF_Calc = SF.CreateScriptService("Calc")
    doc.ClearAll("B2:B7")

# Example of using ClearFormats
def clear_contents_v2(args=None):
    doc: SF.SFDocuments.SF_Calc = SF.CreateScriptService("Calc")
    doc.ClearFormats("D2:D7")

# Example of using ClearValues
def clear_contents_v3(args=None):
    doc: SF.SFDocuments.SF_Calc = SF.CreateScriptService("Calc")
    doc.ClearValues("F2:F7")

# Copying to a single cell
def copy_cells_v1(args=None):
    doc: SF.SFDocuments.SF_Calc = SF.CreateScriptService("Calc")
    doc.CopyToCell("A1:A4", "C1")

# Copying cells into a larger range
def copy_cells_v2(args=None):
    doc: SF.SFDocuments.SF_Calc = SF.CreateScriptService("Calc")
    doc.CopyToRange("A1:A4", "E1:F6")

# Copies range from an open file
def copy_range_from_file(args=None):
    # Reference to current document (destination)
    doc: SF.SFDocuments.SF_Calc = SF.CreateScriptService("Calc")
    # Reference to the source document
    svc: SF.SFScriptForge.SF_UI = SF.CreateScriptService("UI")
    source_doc: SF.SFDocuments.SF_Calc = svc.GetDocument("DataSource.ods")
    source_range = source_doc.Range("Sheet1.A1:A5")
    # Pastes the contents into the destination
    doc.CopyToCell(source_range, "A1")

# Inserting new sheet
def create_sheet_example(args=None):
    doc: SF.SFDocuments.SF_Calc = SF.CreateScriptService("Calc")
    doc.InsertSheet("TestSheet", 2)
    doc.Activate("TestSheet")

# Copying an existing sheet
def copy_sheet_example(args=None):
    doc: SF.SFDocuments.SF_Calc = SF.CreateScriptService("Calc")
    doc.CopySheet("TestSheet", "Copy_TestSheet")

# Removing a sheet
def remove_sheet_example(args=None):
    doc: SF.SFDocuments.SF_Calc = SF.CreateScriptService("Calc")
    doc.RemoveSheet("Copy_TestSheet")

# Copies sheet from another file (open or closed)
def copy_from_file_example(args=None):
    doc: SF.SFDocuments.SF_Calc = SF.CreateScriptService("Calc")
    res_path = get_res_path(doc)
    wb = str(Path(res_path, "DataSource.ods"))
    doc.CopySheetFromFile(wb, "Sheet2", "Copy_Sheet2")

# Example using the DAvg method
def calculate_average(args=None):
    doc: SF.SFDocuments.SF_Calc = SF.CreateScriptService("Calc")
    bas: SF.SFScriptForge.SF_Basic = SF.CreateScriptService("Basic")
    result = doc.DAvg("A1:E1")
    bas.MsgBox("The average is {:.02f}".format(result))

# Open CSV file JobData_v1.csv using default configuration
def open_csv_file_v1(args=None):
    doc: SF.SFDocuments.SF_Calc = SF.CreateScriptService("Calc")
    res_path = get_res_path(doc)
    csvfile = str(Path(res_path, "JobData_v1.csv"))
    doc.ImportFromCSVFile(csvfile, "A1")

# Open CSV file JobData_v2.csv using default configuration
def open_csv_file_v2(args=None):
    doc: SF.SFDocuments.SF_Calc = SF.CreateScriptService("Calc")
    res_path = get_res_path(doc)
    csvfile = str(Path(res_path, "JobData_v2.csv"))
    doc.ImportFromCSVFile(csvfile, "A1")

# Open CSV file using custom configuration
def open_csv_file_v3(args=None):
    doc: SF.SFDocuments.SF_Calc = SF.CreateScriptService("Calc")
    res_path = get_res_path(doc)
    csvfile = str(Path(res_path, "JobData_v2.csv"))
    filter_option = "59,34,UTF-8,1"
    doc.ImportFromCSVFile(csvfile, "A1", filter_option)

# gets the path to res dir
def get_res_path(doc: SF.SFDocuments.SF_Calc) -> Path:
    cmp: XModel = doc.XComponent
    url = cmp.getURL()
    bas: SF.SFScriptForge.SF_Basic = SF.CreateScriptService("Basic")
    file = bas.ConvertFromUrl(url)
    doc_path = Path(file)
    doc_dir = doc_path.parent
    return doc_dir / 'res'
