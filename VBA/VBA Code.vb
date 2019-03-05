                        VBA Notes and Code

Public Sub WriteToA1()

'Writes 100 to cell A1 in ("Sheet1") in MyVBA.xlm
'Note that you have to specify the workbook, sheet and range of cells
'Each are separated by decimals - 

    Workbooks("MyVBA.xlm").Worksheets("Sheet1").Range("A1") = 100

End Sub
-----------------------------------------------------------------------

Public Sub WriteToMulti()

'Uses the same concept as WriteTo__ but allows you to specify what you
'Want to write in multiple cells at the same time

'Writes "John" to cell B1 of wksh "Sheet1" in MyVBA.xlm
    Workbooks("MyVBA.xlm").WorkSheets("Sheet1").Range("B1") = "John"

'Writes 100 to cell A1 of worksheet "Accounts" in MyVBA.xlm
    Workbooks("MyVBA.xlm).Worksheets("Accounts").Range("A1") = 100

'Writes the date to cell D3 of worksheet "Sheet2" in Book.xlsc
    Workbooks("Boook.xlsx").Worksheets("Sheet2").Range("D3") = "1/1/2017"

End Sub
-----------------------------------------------------------------------
                        Using Workbooks

Public Sub WorkbookProperties()

'Prints the number of open workbooks
    Debug.Print Workbooks.Count

'Prints the full workbook name
    Debug.Print Workbooks("Test1.xls").FullName    

'Displays the full workbook name in a message dialogue
    MsgBox Workbooks("Test1.xlsx").FullName

'Prints the number of worksheets in Test2.xlsx
    Debug.Print Workbooks("Test2.xlsx").Worksheets.Count

'Prints the name of the currently active sheet of Test2.xlsx
    Debug.Print Workbooks("Test2.xlsx").ActiveSheet.Name

'Closes workbooked named Test1.xlsx
    Workbooks("Test1.xlsx").Close

'Closes workbook Test2.xlsx and saves changes
    Workbooks("Test2.xlsx").Close SaveChanges:=True

End Sub
-----------------------------------------------------------------------
                        Finding Workbooks

Public Sub PrintWrkFileName()

'This is helpful when you are using loops to specify which one
'Prints out the full filename of all open workbooks

Dim wrk As Workbook
For Each wrk In Workbooks
    Debug.Print wrk.FullName
Next wrk

End Sub
---------------
ALTERNATIVELY.....
---------------
Public Sub PrintWrkFileNameIdx()

'Prints out the full filename of all the open workbooks

Dim i As Long
For i = 1 To Workbooks.Count
    Debug.Print Workbooks(i).FullName
Next i

End Sub
-----------------------------------------------------------------------
                        Using Index - Don't

-----------------------------------------------------------------------
                        Open Workbooks

'The following opens the Workbook "Book1.xlsm" in the "C:\Docs" Folder

Public Sub OpenWrk()

'Opens the WorkBook and prints the number of sheets it contains
    Workbooks.Open ("C:\Docs\Book1.xlsm")

    Debug.Print Workbooks("Book1.xlsm").Worksheets.Count

'Close the workbook without saving
    Workbooks("Book1.xlsm").Close saveChanges:=False

End Sub
-----------------------------------------------------------------------
                        Checking if a Workbook Exists

Public Sub OpenWrkDir()

    If Dir("C:\Docs\Book1.xlsm") = " " Then
	'File does not exist - Inform user
       MsgBox "Could not open the workbook. Please check if it exists"
    Else
	'Open workbook and do something with it
       Workbooks.Open("C:\Docs\Book1.xlsm").Open
    End If

End Sub        
-----------------------------------------------------------------------
                        Checking if a Workbook Is Open
                        If not, open the workbook

Function GetWorkbook(ByVal sFullFilename As String) As Workbook

    Dim sFilename As String
    sFilename = Dir(sFullFilename)

    On Error Resume Next
    Dim wk As Workbook
    Set wk = Workbooks(sFilename)

    If wk Is Nothing Then
       Set wk = Workbooks.Open(sFullFilename)
    End If

    On Error Goto 0
    Set GetWorkbook = wk

End Function

---------
Example
---------

Sub ExampleOpenWorkbook()

    Dim sFilename As String
    sFilename = "C:\Docs\Book2.xlsx"

    Dim wk As Workbook
    Set wk = GetWorkbook(sFilename)

End Sub

---------------------------
If the file is already open
---------------------------

Function IsWorkbookOpen(FileName As String) As Boolean

    Dim ff As Long, ErrNo As Long
'Is this used for an unspecified range?

    On Error Resume Next

    'Open File and store error number
    ff = FreeFile()
    Open FileName For Input Lock Read As #ff
    Close ff
    ErrNo = Err
    On Error Goto 0

-----------------------------------------------------------------------

                        Close Workbook

wk.Close

'Don't save changes when closing
wk.Close ShaveChanges:= False

'Do save changes when closing
wk.Close SaveChanges:= True

'Save without closing
wk.Save

'Save As function
wk.SaveAs "C:\Users\ima....\Example.xlsm

'Save a copy of your workbook
wk.SaveCopyAs "C\Users\ima....\Example.xlsm

'Copy a workbook without opening it

Public Sub CopyWorkbook()
    FileCopy "C:\Docs\Docs.xlsm", "C:\Docs\Example_Copy.xlsm"
End Sub
----------------------------------
'Save the workbook using the file dialogue
Public Function UserSelectWorkbook() as String

    On Error Goto Error Handler

    Dim sWorkbookName As String

Dim FD As FileDialog
Set FD = Application.FileDialog(msoFileDialogFilePicker)

'Open the file dialog
With FD
    'Set Dialog Title
    .Title = "Please Select File"

    'Add Filter
    .Filters.Add "Excel Files", "*.xls;*.xlsx,*.xlsm"

    'Allow selection of one file only
    .AllowMultiSelect = False

    'Display dialog
    .Show

    If FD.SelectedItems.Count > 0 Then
       UserSelectWorkbook = FD.SelectedItems(1)
    Else
       MsgBox "Selecting a file has been cancelled. "
       UserSelectWorkbook = " "
    End If
End With

'Clean up
    Set FD = Nothing
Done:
    Exit Function
ErrorHandler:
    MsgBox "Error: " + Err.Description
End Function    
-----------------------------------------------------------------------
                        Declaring Variables

Public Sub OpenWrkObjects()

    Dim wrk As Workbook
    Set wrk = Workbooks.Open("C:\users\ima....\Book1.xlsm")

    'Print the number of sheets in each book
    Debug.Print wrk.Worksheets.Count
    Debug.Print wrk.Name

    wrk.Close

End Sub
-----------------------------------------------------------------------
                        Create a New Workbook

Public Sub AddWorkbook()

    Dim wrk As Workbook
    Set wrk = Workbooks.Add

    'Save as xlsx - default
    wrk.SaveAs "C:\Temp\Example.xlsx"

    'Save as Macro enabled workbook
    wrk.SaveAs "C:\Temp\Example.xlsm", xlOpenXMLWorkbookMacroEnabled

End Sub
-----------------------------------------------------------------------
                        Using With Keyword

'Not using the With keyword

Public Sub NotUsingWith()

    Debug.Print Workbooks("Book2.xlsm").Worksheets.Count
    Debug.Print Workbooks("Book2.xlsm").Name
    Debug.Print Workbooks("Book2.xlsm")Worksheets(1)Range("A1")
    Workbooks("Book2.xlsm").Close

End Sub

'Using the With keyword to make the code more efficient

Public Sub UsingWith()

    With Workbooks("Book2.xlsm")
       Debug.Print .Worksheets.Count
       Debug.Print .Name
       Debug.Print .Worksheets(1).Range("A1")
       .Close
    End With

End Sub



























