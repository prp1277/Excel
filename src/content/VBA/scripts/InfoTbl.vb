Attribute VB_Name = "InfoTbl"
Option Explicit

Sub InfoTbl()
    Dim NationalAccounts As String
    Dim r As Long
    Dim f As String
    Dim FileSize As Double
    
    With ActiveWorkbook.Sheets.Add
        .Visible = False
        .Name = "Parameters"
        Cells(1, 2) = ThisWorkbook.FullName
    End With
    
    r = 1

    Sheets("Parameter Table").Activate
'   Insert headers
    Cells.ClearContents
    Cells(r, 1) = ThisWorkbook.FullName
    Cells(r, 2) = "Files in " & NationalAccounts
    Cells(r, 3) = "Size"
    Cells(r, 4) = "Date/Time"
    Range("A1:C1").Font.Bold = True
    
'   Get first file
    f = Dir(NationalAccounts, vbReadOnly + vbHidden + vbSystem)
    Do While f <> ""
        r = r + 1
        Cells(r, 1) = f
        'adjust for filesize > 2 gigabytes
        FileSize = FileLen(NationalAccounts & f)
        If FileSize < 0 Then FileSize = FileSize + 4294967296#
        Cells(r, 2) = FileSize
        Cells(r, 3) = FileDateTime(NationalAccounts & f)
    '   Get next file
        f = Dir
    Loop
    
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$1:$C$8"), , xlYes).Name = "ParameterTable"

End Sub

Sub PrintInfo()

Debug.Print ThisWorkbook.path
Debug.Print ThisWorkbook.FullName


End Sub
