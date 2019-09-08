Attribute VB_Name = "Shortcuts"
Option Explicit

Sub ZapContents()
Attribute ZapContents.VB_Description = "Clear contents from cells "
Attribute ZapContents.VB_ProcData.VB_Invoke_Func = "Q\n14"
'ZapContents Macro
'Clears the contents using active cell, C+S+Right & C+S+Down
    Application.ScreenUpdating = False
        
    ActiveCell.Activate
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.ClearContents
    Cells.Select
    Selection.Columns.AutoFit
End Sub

Sub QuickFormat()
Attribute QuickFormat.VB_Description = "Autofit, align left and clear wrapped and merged cells"
Attribute QuickFormat.VB_ProcData.VB_Invoke_Func = "M\n14"
'Clear Formatting from all cells
    Application.ScreenUpdating = False
    
ThisWorkbook.Activate
Cells.ClearFormats

    With Cells
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        .Columns.AutoFit
    End With
End Sub

Sub ActiveCellToTable()
Attribute ActiveCellToTable.VB_Description = "Select active cell to last cell and convert to table"
Attribute ActiveCellToTable.VB_ProcData.VB_Invoke_Func = "E\n14"
' Keyboard Shortcut: Ctrl+Shift+E
    Application.ScreenUpdating = False
    
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes).Name = _
        "Table1"
End Sub

Sub ConvertToHyperlink()
Attribute ConvertToHyperlink.VB_Description = "Convert Active Cell to Hyperlink"
Attribute ConvertToHyperlink.VB_ProcData.VB_Invoke_Func = "H\n14"
'Keyboard Shortcut Ctrl + Shift + H
    
    ActiveSheet.Hyperlinks.Add Anchor:=Excel.Selection, Address:= _
        ActiveCell.Value, TextToDisplay:=ActiveCell.Value
End Sub

Private Sub Workbook_BeforeClose()
    Dim FP As String
    FP = "C:\Users\imami\OneDrive\Documents\Shared\Templates"
    Dim FN As String
    FN = "\Personal.xlam"
    
    Windows("PERSONAL.XLSB").Visible = False
    ThisWorkbook.Save
    Workbooks(1).SaveCopyAs (FP & FN)
    '(("C:\Users\imami\OneDrive\Documents\Shared\PERSONAL") & ".xlsm")
End Sub
