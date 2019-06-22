Sub WorkGroups()

    'Instantiate variables
    Dim maxGroup As Integer
    Dim dropSheets As Integer
    Dim ws As Worksheet
    Dim i As Integer
    Dim targetGroup As Integer
    Dim targetRow As Integer
    Dim name As String

    'Determine the number of groups
    maxGroup = WorksheetFunction.Max(Columns(2))

    'Drop group worksheets
    dropSheets = MsgBox("Do you want to delete the current Group Worksheets? ", vbYesNo, "Delete Worksheets")
    If dropSheets = vbYes Then
        Application.DisplayAlerts = False
        For Each ws In ActiveWorkbook.Worksheets
            If ws.name <> "Master" Then
                ws.Delete
            End If
        Next
        Application.DisplayAlerts = True

        'Create new group worksheets
        For i = maxGroup To 1 Step -1
            Set ws = Worksheets.Add(After:=Worksheets("Master"))
            ws.name = "Group " & i
            ws.Cells(1, 1).Value = "Name"
            ws.Cells(1, 1).Font.Bold = True
        Next i

        'Remove Copied notation
        i = 2
        Do
            If Worksheets("Master").Cells(i, 3).Value = "Copied" Then
                Worksheets("Master").Cells(i, 3).Value = ""
            End If
            i = i + 1
        Loop Until Worksheets("Master").Cells(i, 2).Value = ""

    End If

    'Copy values to respective group worksheets
    i = 2
    Do
        If Worksheets("Master").Cells(i, 3).Value <> "Copied" Then
            targetGroup = Worksheets("Master").Cells(i, 2).Value
            name = Worksheets("Master").Cells(i, 1).Value

            'Determine next available row in group worksheet
            targetRow = 2
            Do Until Worksheets("Group " & targetGroup).Cells(targetRow, 1).Value = ""
                targetRow = targetRow + 1
            Loop

            Worksheets("Group " & targetGroup).Cells(targetRow, 1).Value = name
            Worksheets("Master").Cells(i, 3).Value = "Copied"
        End If
        i = i + 1
    Loop Until Worksheets("Master").Cells(i, 2).Value = ""

    Worksheets("Master").Activate
End Sub