Option Explicit

Sub InfoTbl()

Dim Wb As Workbook
Dim Sht As Worksheet
Set Wb = ThisWorkbook
Set Sht = ThisWorkbook.Worksheets(1)

'Activate leftmost worksheet and clear the cells 
Sht.Activate
Cells.ClearContents

' RxC Notation - Column 1, Rows 1 - 5
Cells(1, 1) = "File Name"
Cells(2, 1) = "FilePath"
Cells(3, 1) = "Folder"
Cells(4, 1) = "Sheet Index"
Cells(5, 1) = "Last Updated"

' RxC Notation - Column 2, Rows 1 - 5
Cells(1, 2) = Wb.Name ' Filename
Cells(2, 2) = Wb.FullName
Cells(3, 2) = MsoFileDialogView.msoFileDialogViewList
Cells(4, 2) = "=SHEET(A1)"
Cells(5, 2) = "=Now()"

End Sub