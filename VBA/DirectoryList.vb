Option Explicit

Public FSO As New FileSystemObject
Private FileType As Variant
'------------------------------------------------------------------------------

Sub ListHyperlinkFilesInSubFolders()

    ' Written by Philip Treacy, http://www.myonlinetraininghub.com/author/philipt
    ' My Online Training Hub http://www.myonlinetraininghub.com/Create-Hyperlinked-List-of-Files-in-Subfolders
    ' May 2014

    Dim StartingCell As String 'Cell where hyperlinked list starts
    Dim FSOFolder As Folder
    Dim RootFolder As String

    Application.ScreenUpdating = False
    
    'Make this a cell address to insert list at fixed cell
    'e.g. StartingCell = "A1"
    StartingCell = ActiveCell.Address


    'Ask for folder to list files from
    With Application.FileDialog(msoFileDialogFolderPicker)
    
        .InitialFileName = Application.DefaultFilePath & "\"
        .Title = "Please select folder to list files from"
        .Show
    
        'If a folder has been selected
        If .SelectedItems.Count <> 0 Then
        
            RootFolder = .SelectedItems(1)
            
            Set FSOFolder = FSO.GetFolder(RootFolder)
            
            'Ask what type of files to look for
            FileType = Application.InputBox("* and ? wildcards are valid " & vbCrLf & vbCrLf & " e.g. .xls* to list XLS, XLSX and XLSM" _
                        & vbCrLf & vbCrLf & "??st.* to list West.xlsx and East.xlsx" & vbCrLf & vbCrLf & "Just click OK to list all files.", _
                        "What type of files do you want to list?", "")
                        
            If FileType = False Then 'Cancel pressed
                
                MsgBox "Process Cancelled"
                Exit Sub

            ElseIf FileType = vbNullString Then 'Nothing entered and OK pressed

                FileType = "*.*"
            
            End If
            
            'Clear the active sheet to remove previous results
            ActiveSheet.Cells.Clear

            'Enter default message in case no files are in folder
            With Range(StartingCell)
            
                .ClearFormats
                .Value = "No " & FileType & " files found in " & RootFolder
                .Select
                
            End With
            
            ' Call recursive sub to list files
            ListFilesInSubFolders FSOFolder, ActiveCell
    
            'Autofit the columns containing our results
            Columns.AutoFit
            
        Else
        
            'If no folder selected, admonish user for wasting CPU cycles :)
            MsgBox "No folder selected.", vbExclamation
        
        End If

    End With
    
    Application.ScreenUpdating = True

End Sub
'------------------------------------------------------------------------------


Sub ListFilesInSubFolders(StartingFolder As Scripting.Folder, DestinationRange As Range)
    ' Written by Philip Treacy, http://www.myonlinetraininghub.com/author/philipt
    ' My Online Training Hub http://www.myonlinetraininghub.com/Create-Hyperlinked-List-of-Files-in-Subfolders
    ' May 2014
    ' Lists all files specified by FileType in all subfolders of the StartingFolder object.
    ' This sub is called recursively
    
    Dim CurrentFilename As String
    Dim OffsetRow As Long
    Dim TargetFiles As String
    Dim SubFolder As Scripting.Folder
    
    'Write name of folder to cell
    DestinationRange.Value = StartingFolder.Path
    
    'Get the first file, look for Normal, Read Only, System and Hidden files
    TargetFiles = StartingFolder.Path & "\" & FileType
                
            CurrentFilename = Dir(TargetFiles, 7)
            
            OffsetRow = 1
            
            Do While CurrentFilename <> ""
            
                'Create the hyperlink
                DestinationRange.Offset(OffsetRow).Hyperlinks.Add Anchor:=DestinationRange.Offset(OffsetRow), Address:=StartingFolder.Path & "\" & CurrentFilename, TextToDisplay:=CurrentFilename
                
                OffsetRow = OffsetRow + 1

                'Get the next file
                CurrentFilename = Dir
        
            Loop


    ' Offset the DestinationRange one column to the right and OffsetRows down so that we start listing files
    ' inthe next folder below where we just finished. This results in an indented view of the folder structure
    Set DestinationRange = DestinationRange.Offset(OffsetRow)
    
    ' For each SubFolder in the current StartingFolder call ListFilesInSubFolders (recursive)
    ' The sub continues to call itself for each and every folder it finds until it has
    ' traversed all folders below the original StartingFolder
    For Each SubFolder In StartingFolder.SubFolders
        
        ListFilesInSubFolders SubFolder, DestinationRange
        
    Next SubFolder
    
    ' Once all files in SubFolder are listed, move the DestinationRange down 1 row and left 1 column.
    ' This gives a clear visual structure to the listing showing that we are done with the current SubFolder
    ' and moving on to the next SubFolder
    'Set DestinationRange = DestinationRange.Offset(1)
    'DestinationRange.Select
    
    End Sub

