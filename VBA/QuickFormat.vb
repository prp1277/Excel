Sub QuickFormat()
'Clear Formatting from all cells
    Application.ScreenUpdating = False
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