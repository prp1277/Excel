Attribute VB_Name = "Functions"
Option Explicit

Private Function GetValue(path, file, sheet, ref)
'   Retrieves a value from a closed workbook
    Dim arg As String

'   Make sure the file exists
    If Right(path, 1) <> "\" Then path = path & "\"
    If Dir(path & file) = "" Then
        GetValue = "File Not Found"
        Exit Function
    End If

'   Create the argument
    arg = "'" & path & "[" & file & "]" & sheet & "'!" & _
      Range(ref).Range("A1").Address(, , xlR1C1)

'   Execute an XLM macro
    GetValue = ExecuteExcel4Macro(arg)
End Function

Sub TestGetValue()
    Dim p As String, f As String
    Dim s As String, a As String
    
    p = ThisWorkbook.path
    f = ThisWorkbook.FullName
    s = "Sheet2"
    a = "C1"
    MsgBox GetValue(p, f, s, a)
End Sub

Sub TestGetValue2()
    Dim p As String, f As String
    Dim s As String, a As String
    Dim r As Long, c As Long
   
    p = ThisWorkbook.path
    f = "myworkbook.xlsx"
    s = "Sheet1"
    Application.ScreenUpdating = False
    For r = 1 To 100
        For c = 1 To 12
            a = Cells(r, c).Address
            Cells(r, c) = GetValue(p, f, s, a)
        Next c
    Next r
End Sub

Private Function GetValueByOpeningTheFile(path, file, sheet, ref)
    Dim wb As Workbook
    Application.ScreenUpdating = False
    Set wb = Workbooks.Open(path & Application.PathSeparator & file)
    Worksheets(sheet).Activate
    GetValueByOpeningTheFile = Range(ref)
    wb.Close savechanges:=False
    Application.ScreenUpdating = True
End Function

Sub TestGetValueByOpeningTheFile()
    Dim p As String, f As String
    Dim s As String, a As String
    
    p = ThisWorkbook.path
    f = "myworkbook.xlsx"
    s = "Sheet2"
    a = "C1"
    MsgBox GetValueByOpeningTheFile(p, f, s, a)
End Sub

Private Function FileExists(fname) As Boolean
'   Returns TRUE if the file exists
    Dim x As String
    x = Dir(fname)
    If x <> "" Then FileExists = True _
        Else FileExists = False
End Function

Private Function FileNameOnly(pname) As String
'   Returns the filename from a path/filename string
    Dim temp As Variant
    Length = Len(pname)
    temp = Split(pname, Application.PathSeparator)
    FileNameOnly = temp(UBound(temp))
End Function

Private Function PathExists(pname) As Boolean
'   Returns TRUE if the path exists
  If Dir(pname, vbDirectory) = "" Then
    PathExists = False
 Else
    PathExists = (GetAttr(pname) And vbDirectory) = vbDirectory
 End If
End Function

Private Function RangeNameExists(nname) As Boolean
'   Returns TRUE if the range name exists
    Dim n As Name
    RangeNameExists = False
    For Each n In ActiveWorkbook.Names
        If UCase(n.Name) = UCase(nname) Then
            RangeNameExists = True
            Exit Function
        End If
    Next n
End Function

Private Function SheetExists(sname) As Boolean
'   Returns TRUE if sheet exists in the active workbook
    Dim x As Object
    On Error Resume Next
    Set x = ActiveWorkbook.Sheets(sname)
    If Err = 0 Then SheetExists = True _
        Else SheetExists = False
End Function

Private Function WorkbookIsOpen(wbname) As Boolean
'   Returns TRUE if the workbook is open
    Dim x As Workbook
    On Error Resume Next
    Set x = Workbooks(wbname)
    If Err = 0 Then WorkbookIsOpen = True _
        Else WorkbookIsOpen = False
End Function

Private Function IsInCollection(Coln As Object, Item As String) As Boolean
    Dim Obj As Object
    On Error Resume Next
    Set Obj = Coln(Item)
    IsInCollection = Not Obj Is Nothing
End Function
