Option Explicit
'This proceedure opens the files from the links, saves 
'them into the folder and cleans them to be ready for
'the query editor to append them together
Sub FollowLink()
'
' FollowLink Macro
Application.DisplayAlerts = False
Application.ScreenUpdating = False

Dim Wb As Workbook, Wbs As Workbooks
Dim Sh As Worksheet, Shs As Sheets, j As Integer
Dim FP As String, FN As String, QF As String
Dim i As Integer, DFP As String, Tbl As ListObject

'Set Variables
Set Tbl = Sheets("Links&Locations").ListObjects("LinkSiteSearch")
DFP = Application.DefaultFilePath
QF = "C:\Users\prp12.000\OneDrive\prp1277.github.io\_data\[XlsFiles]VBATarget\"
FP = "C:\Users\prp12.000\OneDrive\prp1277.github.io\_data\QueryFolderTarget\"

'Activate, count how many workbooks are open and print default file location
Sheets("Links&Locations").Activate
j = Workbooks.Count
Cells(1, 1) = DFP

'These 4 lines loop through the table
'And Open each link in column 1
For i = 1 To Tbl.Range.Rows.Count
    Tbl.Range.Columns(1).Select
    Selection.Hyperlinks(1).Follow NewWindow:=False, AddHistory:=True
    Workbooks("JobCleanXLS.xlsm").Sheets("Links&Locations").Activate
Next i
    
'Now, all the workbooks are opened,
'clean them and save them to the query folder
For i = 1 To (Workbooks.Count - j)
    'Clear the shitty format
    Application.Run "JobCleanXls.xlsm!QuickFormat"
    'Delete the stupid picture
    ActiveSheet.Shapes.Range(Array("Picture 1")).Delete
    'Why's there a blank column in the first place?
    ActiveCell.Columns("A:A").EntireColumn.Delete Shift:=xlToLeft
    'Let alone three blank rows
    ActiveCell.Rows("1:3").EntireRow.Delete Shift:=xlUp
    'Delete Hyperlinks because Power Query doesn't see them
    Cells.Hyperlinks.Delete
    'Freezing panes is retarded
    ActiveWindow.FreezePanes = False
    'Close the window cause we're done with that shit
    ActiveWindow.Close
Next i
    
Workbooks("JobCleanXLS.xlsm").Activate
End Sub

-
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
