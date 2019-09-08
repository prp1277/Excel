Attribute VB_Name = "HideOnClose"
Option Explicit

Private Sub Workbook_BeforeClose()
'OneDrive Sharable link ("https://1drv.ms/f/s!AoLkNBOSNnKygZp2OTAaUkzYjQ_5_A")

Dim MWb As Workbook
Set MWb = Workbooks("Personal.xlsb")
Dim HFileName As String, WFileName As String
HFileName = "C:\Users\imami\OneDrive\Documents\Shared\Templates\PERSONAL.xlam"
WFileName = "C:\Users\PRPowell\OneNote\Shared\Templates\PERSONAL.xlsb"


    Debug.Print HFileName, WFileName
    Windows("PERSONAL.XLSB").Visible = False
    MWb.Save
    MWb.SaveCopyAs Filename:=HFileName
    'MWb.SaveCopyAs Filename:=WFileName
    
End Sub
