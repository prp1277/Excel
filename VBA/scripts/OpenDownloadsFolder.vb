1/1/18 - Open from Downloads, delete rows, and format as table

Option Explicit

Sub Macro1()
'
' Macro1 Macro
'

'
    ChDir "C:\Users\imami\Downloads"
    Workbooks.Open Filename:="C:\Users\imami\Downloads\transactions (15).csv"
    Application.Run "PERSONAL.XLSB!QuickFormat"
    Rows("1:12").Select
    Selection.Delete Shift:=xlUp
    Range("A1").Select
    Application.Run "PERSONAL.XLSB!ActiveCellToTable"
    ActiveSheet.ListObjects("Table1").Name = "NewTable"
End Sub
---------------------------------------------------------------------------------------
Sub Macro2()
'
' Macro2 Macro
'

'
    ChDir "C:\Users\imami\OneDrive\Documents\Resources\Excel"
    ActiveWorkbook.SaveAs Filename:= _
        "https://d.docs.live.net/b27236921334e482/Documents/Resources/Excel/MacroWorkbook.xlsm" _
        , FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
End Sub
