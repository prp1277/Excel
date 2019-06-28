Option Explicit
'------------------------------------------------------------------------------
Sub ZapContents()
'ZapContents Macro
'Clears the contents using active cell, C+S+Right & C+S+Down
    Application.ScreenUpdating = False
        
    ActiveCell.Activate
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.ClearContents
    Cells.Select
    Selection.Columns.AutoFit

    Application.ScreenUpdating = True
End Sub
'------------------------------------------------------------------------------
Sub QuickFormat()
'Clear Formatting from all cells
'ctrl+shift+m
    Application.ScreenUpdating = False

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

    Application.ScreenUpdating = True
End Sub
'------------------------------------------------------------------------------
Sub ActiveCellToTable()
' Keyboard Shortcut: Ctrl+Shift+E
    Application.ScreenUpdating = False
    
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes).Name = _
        "Table1"

    Application.ScreenUpdating = True
End Sub
'------------------------------------------------------------------------------
Sub ConvertToHyperlink()
     'Keyboard Shortcut Ctrl + Shift + H
    
    ActiveSheet.Hyperlinks.Add Anchor:=Excel.Selection, Address:= _
        ActiveCell.Value, TextToDisplay:=ActiveCell.Value

End Sub
'------------------------------------------------------------------------------
Sub RemoveHyperlinks()
With Cells
    .ClearHyperlinks
    .ClearFormats
    .Font.Color = Default
    End With
    
Range("A1").Select
Selection.Columns.AutoFit

End Sub
'------------------------------------------------------------------------------
Sub PasteLinkRight()
' This macro works as if you used ctrl + r
' Except it strips the hyperlink and pastes as text
' Ctrl + Shift + R

Dim Lnk As String
Lnk = ActiveCell.Offset(0, -1).Hyperlinks(1).Address
    With ActiveCell
        .Activate
        .Value = Lnk
        .Offset(0, -1).Hyperlinks(1).Delete
    End With

End Sub
