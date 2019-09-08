Option Explicit

Sub QueryIndex()
'This macro loops through the workbooks' queries
'and adds a Query Index page with the M formula
Dim i As Integer
Dim ws As Worksheet
Dim QCount As Integer
Dim QName As String
Dim QForm As String

Set ws = Worksheets.Add(Before:=Worksheets(1))

QCount = ActiveWorkbook.Queries.Count

ws.name = "Query Index"
ws.Cells(1, 1).Value = "Name"
ws.Cells(1, 2).Value = "Query String"

For i = 1 To QCount
  QName = ActiveWorkbook.Queries.Item(i).name
  QForm = ActiveWorkbook.Queries.Item(i).Formula
  ws.Cells(i + 1, 1).Value = QName
  ws.Cells(i + 1, 2).Value = QForm
Next i

With ActiveSheet.Cells
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
    With ActiveSheet.Cells.Font
        .name = "Calibri"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With

End Sub
