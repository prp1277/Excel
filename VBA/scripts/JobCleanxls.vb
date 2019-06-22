Sub ConvertToXlsx()
Dim FP As String, FN As String, QF As String
Dim TWB As Workbook, Wb As Workbooks
Dim Sh As Sheets
Dim DFP As String

Application.ScreenUpdating = False
Application.DisplayAlerts = False

Set TWB = Workbooks("JobCleanXls1.xlsm")

'Replace "C:\...." with where you saved the xls files
QF = "C:\Users\prp12.000\OneDrive\prp1277.github.io\_data\[XlsFiles]VBATarget\"

'Replace "C:\..." with where you want to save the xlsx files
FP = "C:\Users\prp12.000\OneDrive\prp1277.github.io\_data\QueryFolderTarget\"


'1 - Career Builder Analyst File
With Application.Workbooks.Open(QF & "CBAnalyst.xls")
    Application.Run "JobCleanXls.xlsm!QuickFormat"
    ActiveSheet.Shapes.Range(Array("Picture 1")).Delete
    ActiveCell.Columns("A:A").EntireColumn.Delete Shift:=xlToLeft
    ActiveCell.Rows("1:3").EntireRow.Delete Shift:=xlUp
    Cells.Hyperlinks.Delete
    ActiveWindow.FreezePanes = False
    ActiveWorkbook.SaveAs FP & "CBAnalyst.xlsx" _
    , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    ActiveWindow.Close
End With

'2 - Career Builder Data File
With Application.Workbooks.Open(QF & "CBData.xls")
    Application.Run "JobCleanXls.xlsm!QuickFormat"
    ActiveSheet.Shapes.Range(Array("Picture 1")).Delete
    ActiveCell.Columns("A:A").EntireColumn.Delete Shift:=xlToLeft
    ActiveCell.Rows("1:3").EntireRow.Delete Shift:=xlUp
    Cells.Hyperlinks.Delete
    ActiveWindow.FreezePanes = False
    ActiveWorkbook.SaveAs FP & "CBData.xlsx" _
    , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    ActiveWindow.Close
End With

'3 - Career Builder Excel File
With Application.Workbooks.Open(QF & "CBExcel.xls")
    Application.Run "JobCleanXls.xlsm!QuickFormat"
    ActiveSheet.Shapes.Range(Array("Picture 1")).Delete
    ActiveCell.Columns("A:A").EntireColumn.Delete Shift:=xlToLeft
    ActiveCell.Rows("1:3").EntireRow.Delete Shift:=xlUp
    Cells.Hyperlinks.Delete
    ActiveWindow.FreezePanes = False
    ActiveWorkbook.SaveAs FP & "CBExcel.xlsx" _
    , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    ActiveWindow.Close
End With

'4 - Career Builder Financial Analyst File
With Application.Workbooks.Open(QF & "CBFinancialAnalyst.xls")
    Application.Run "JobCleanXls.xlsm!QuickFormat"
    ActiveSheet.Shapes.Range(Array("Picture 1")).Delete
    ActiveCell.Columns("A:A").EntireColumn.Delete Shift:=xlToLeft
    ActiveCell.Rows("1:3").EntireRow.Delete Shift:=xlUp
    Cells.Hyperlinks.Delete
    ActiveWindow.FreezePanes = False
    ActiveWorkbook.SaveAs FP & "CBFinancialAnalyst.xlsx" _
    , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    ActiveWindow.Close
End With

'5 - Indeed Analyst File
With Application.Workbooks.Open(QF & "INAnalyst.xls")
    Application.Run "JobCleanXls.xlsm!QuickFormat"
    ActiveSheet.Shapes.Range(Array("Picture 1")).Delete
    ActiveCell.Columns("A:A").EntireColumn.Delete Shift:=xlToLeft
    ActiveCell.Rows("1:3").EntireRow.Delete Shift:=xlUp
    Cells.Hyperlinks.Delete
    ActiveWindow.FreezePanes = False
    ActiveWorkbook.SaveAs FP & "INAnalyst.xlsx" _
    , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    ActiveWindow.Close
End With

'6 - Indeed Data File
With Application.Workbooks.Open(QF & "INData.xls")
    Application.Run "JobCleanXls.xlsm!QuickFormat"
    ActiveSheet.Shapes.Range(Array("Picture 1")).Delete
    ActiveCell.Columns("A:A").EntireColumn.Delete Shift:=xlToLeft
    ActiveCell.Rows("1:3").EntireRow.Delete Shift:=xlUp
    Cells.Hyperlinks.Delete
    ActiveWindow.FreezePanes = False
    ActiveWorkbook.SaveAs FP & "INData.xlsx" _
    , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    ActiveWindow.Close
End With

'7 - Indeed Excel File
With Application.Workbooks.Open(QF & "INExcel.xls")
    Application.Run "JobCleanXls.xlsm!QuickFormat"
    ActiveSheet.Shapes.Range(Array("Picture 1")).Delete
    ActiveCell.Columns("A:A").EntireColumn.Delete Shift:=xlToLeft
    ActiveCell.Rows("1:3").EntireRow.Delete Shift:=xlUp
    Cells.Hyperlinks.Delete
    ActiveWindow.FreezePanes = False
    ActiveWorkbook.SaveAs FP & "INExcel.xlsx" _
    , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    ActiveWindow.Close
End With

'8 - Indeed Financial Analyst File
With Application.Workbooks.Open(QF & "INFinancialAnalyst.xls")
    Application.Run "JobCleanXls.xlsm!QuickFormat"
    ActiveSheet.Shapes.Range(Array("Picture 1")).Delete
    ActiveCell.Columns("A:A").EntireColumn.Delete Shift:=xlToLeft
    ActiveCell.Rows("1:3").EntireRow.Delete Shift:=xlUp
    Cells.Hyperlinks.Delete
    ActiveWindow.FreezePanes = False
    ActiveWorkbook.SaveAs FP & "INFinancialAnalyst.xlsx" _
    , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    ActiveWindow.Close
End With

'9 - US Jobs Analyst File
With Application.Workbooks.Open(QF & "USJAnalyst.xls")
    Application.Run "JobCleanXls.xlsm!QuickFormat"
    ActiveSheet.Shapes.Range(Array("Picture 1")).Delete
    ActiveCell.Columns("A:A").EntireColumn.Delete Shift:=xlToLeft
    ActiveCell.Columns("F:F").EntireColumn.Delete Shift:=xlToLeft
    ActiveCell.Rows("1:3").EntireRow.Delete Shift:=xlUp
    Cells.Hyperlinks.Delete
    ActiveWindow.FreezePanes = False
    ActiveWorkbook.SaveAs FP & "USJAnalyst.xlsx" _
    , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    ActiveWindow.Close
End With

'10 - US Jobs Data File
With Application.Workbooks.Open(QF & "USJData.xls")
    Application.Run "JobCleanXls.xlsm!QuickFormat"
    ActiveSheet.Shapes.Range(Array("Picture 1")).Delete
    ActiveCell.Columns("A:A").EntireColumn.Delete Shift:=xlToLeft
    ActiveCell.Columns("F:F").EntireColumn.Delete Shift:=xlToLeft
    ActiveCell.Rows("1:3").EntireRow.Delete Shift:=xlUp
    Cells.Hyperlinks.Delete
    ActiveWindow.FreezePanes = False
    ActiveWorkbook.SaveAs FP & "USJData.xlsx" _
    , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    ActiveWindow.Close
End With

'11 - US Jobs Excel File
With Application.Workbooks.Open(QF & "USJExcel.xls")
    Application.Run "JobCleanXls.xlsm!QuickFormat"
    ActiveSheet.Shapes.Range(Array("Picture 1")).Delete
    ActiveCell.Columns("A:A").EntireColumn.Delete Shift:=xlToLeft
    ActiveCell.Columns("F:F").EntireColumn.Delete Shift:=xlToLeft
    ActiveCell.Rows("1:3").EntireRow.Delete Shift:=xlUp
    Cells.Hyperlinks.Delete
    ActiveWindow.FreezePanes = False
    ActiveWorkbook.SaveAs FP & "USJExcel.xlsx" _
    , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    ActiveWindow.Close
End With

'12 - US Jobs Financial Analyst File
With Application.Workbooks.Open(QF & "USJFinancialAnalyst.xls")
    Application.Run "JobCleanXls.xlsm!QuickFormat"
    ActiveSheet.Shapes.Range(Array("Picture 1")).Delete
    ActiveCell.Columns("A:A").EntireColumn.Delete Shift:=xlToLeft
    ActiveCell.Columns("F:F").EntireColumn.Delete Shift:=xlToLeft
    ActiveCell.Rows("1:3").EntireRow.Delete Shift:=xlUp
    Cells.Hyperlinks.Delete
    ActiveWindow.FreezePanes = False
    ActiveWorkbook.SaveAs FP & "USJFinancialAnalyst.xlsx" _
    , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    ActiveWindow.Close
End With

TWB.Activate

End Sub

'---

Sub FollowLink()
'
' FollowLink Macro
Dim Wb As Workbook, Wbs As Workbooks
Dim Sh As Worksheet, Shs As Sheets
Dim i As Integer, DFP As String, Tbl As String

Tbl = ActiveSheet.ListObjects("LinkSiteSearch")
DFP = Application.DefaultFilePath

Cells(1, 1) = DFP

'These 4 lines loop through the table
'And Open each link in column 1
For i = 1 To Tbl.Range.Rows.Count
    Tbl.Range.Columns(1).Select
    Selection.Hyperlinks(1).Follow NewWindow:=False, _
    AddHistory:=True
Next i

End Sub

'---

Sub QuickFormat()
'Clear Formatting from all cells
    Application.ScreenUpdating = False
    
Debug.Print Workbooks(Workbooks.Count).Name
    Workbooks(Workbooks.Count).Activate

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