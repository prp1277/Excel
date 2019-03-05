Sub PivotData()
'
' PivotFormat Macro
' Formats the SQL Export into Pivot-Ready data
'

'
    Sheets("Invoices").Select
    Range("A5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.TextToColumns Destination:=Range("B6"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 9), Array(5, 1)), TrailingMinusNumbers:=True
    Range("B5").Select
    Sheets.Add
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Invoices!R5C1:R1000C17", Version:=xlPivotTableVersion15).CreatePivotTable _
        TableDestination:="Sheet1!R3C1", TableName:="ALLERRORS", DefaultVersion _
        :=xlPivotTableVersion15
    Sheets("Sheet2").Select
    Sheets("Sheet2").Name = "ALL ERRORS"
    Sheets("Macro").Select
    ActiveWorkbook.SaveCopyAs ("C:\Users\PRPowell\Desktop\Nonshared\Perfect Orders\Kansas City\" & Format("Perfect Order - KC mm.dd.yyy") & ".xlsm")
    ActiveWorkbook.Close SaveChanges:=False
    Workbooks.Open ("C:\Users\PRPowell\Desktop\Nonshared\Perfect Orders\Kansas City\" & ("Perfect Order - KC mm.dd.yyy"))
End Sub

-----------------------------------------------------------------------------------------------
Sub ALLERRORS()
'
' ALL ERRORS Macro
' Creates the ALLERRORS Table
'

'
    Sheets("ALL ERRORS").Select
    With ActiveSheet.PivotTables("ALLERRORS").PivotFields("A#")
        .Orientation = xlRowField
        .Position = 1
    With ActiveSheet.PivotTables("ALLERRORS").PivotFields("A#")
        .PivotItems("").Visible = False
    End With
    End With
    With ActiveSheet.PivotTables("ALLERRORS").PivotFields("Customer")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("ALLERRORS").PivotFields("Invoice Date")
        .Orientation = xlRowField
        .Position = 3
    End With
    With ActiveSheet.PivotTables("ALLERRORS").PivotFields("Invoice  #")
        .Orientation = xlRowField
        .Position = 4
    End With
    With ActiveSheet.PivotTables("ALLERRORS").PivotFields("L1 Error")
        .Orientation = xlRowField
        .Position = 5
    End With
    ActiveWindow.SmallScroll Down:=3
    With ActiveSheet.PivotTables("ALLERRORS").PivotFields("Responsible")
        .Orientation = xlRowField
        .Position = 6
    End With
    With ActiveSheet.PivotTables("ALLERRORS").PivotFields("L1 Error")
        .Orientation = xlColumnField
        .Position = 1
    End With
    ActiveSheet.PivotTables("ALLERRORS").AddDataField ActiveSheet.PivotTables( _
        "ALLERRORS").PivotFields("L1 Error"), "Count of L1 Error", xlCount
    With ActiveSheet.PivotTables("ALLERRORS").PivotFields("L1 Error")
        .Orientation = xlColumnField
        .Position = 1
    End With
    Range("A4").Select
    ActiveSheet.PivotTables("ALLERRORS").Name = "ALLERRORS"
    With ActiveSheet.PivotTables("ALLERRORS")
        .InGridDropZones = True
        .RowAxisLayout xlTabularRow
    End With
    ActiveSheet.PivotTables("ALLERRORS").PivotFields("Customer").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("ALLERRORS").PivotFields("Invoice Date"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("ALLERRORS").PivotFields("Invoice  #").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("ALLERRORS").PivotFields("Responsible").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    With ActiveSheet.PivotTables("ALLERRORS").PivotFields("Invoice  #")
        .Orientation = xlRowField
        .Position = 3
    Sheets("Macro").Select
    End With
End Sub


-----------------------------------------------------------------------------------------------
Sub ErrorPerCustomer()
'
' ErrorPerCust Macro
' Creates the Error/Customer Tab
'

'
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet3").Select
    Sheets("Sheet3").Name = "# Issues By Customer"
    ActiveCell.FormulaR1C1 = "A#"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "LOOKUP VALUE"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Customer"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Total"
    Range("B2").Select
    Sheets("Invoices").Select
    Range("A704").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A6:A704").Select
    Range("A704").Activate
    Selection.Copy
    Sheets("# Issues By Customer").Select
    Range("A2").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveSheet.Range("$A$1:$A$700").RemoveDuplicates Columns:=1, Header:=xlYes
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "=CONCATENATE(RC[-1],"" "",R1C4)"
    Range("B2").Select
    Selection.AutoFill Destination:=Range("B2:B329")
    Range("B2:B329").Select
    Range("C2").Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-1],'ALLERRORS'!R[1]C[-2]:R[538]C[6],2,FALSE)"
    Range("C2").Select
    Selection.AutoFill Destination:=Range("C2:C329")
    Range("C2:C329").Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-2],'ALLERRORS'!R[1]C[-2]:R[538]C[6],2,FALSE)"
    Range("C2").Select
    Selection.AutoFill Destination:=Range("C2:C329")
    Range("C2:C329").Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-2],'ALLERRORS'!R3C1:R540C9,2,FALSE)"
    Range("C2").Select
    Selection.AutoFill Destination:=Range("C2:C329")
    Range("C2:C329").Select
    Range("D2").Select
    Sheets("ALLERRORS").Select
    With ActiveSheet.PivotTables("ALLERRORS").PivotFields("A#")
        .Orientation = xlRowField
        .Position = 1
    End With
    Sheets("# Issues By Customer").Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-2],'ALLERRORS'!R3C1:R540C9,9,FALSE)"
    Range("D2").Select
    Selection.AutoFill Destination:=Range("D2:D329")
    Range("D2:D329").Select

End Sub


-----------------------------------------------------------------------------------------------
Sub ExecutionErrors()
'
' ExecutionErrorPivot Macro
' Creates the Execution Error Pivot Table
'

'
    Sheets.Add
    ActiveWorkbook.Worksheets("ALLERRORS").PivotTables("PlatinumPivot"). _
        PivotCache.CreatePivotTable TableDestination:="Sheet5!R3C1", TableName:= _
        "PivotTable2", DefaultVersion:=xlPivotTableVersion15
    Sheets("Sheet5").Select
    Cells(3, 1).Select
    Sheets("Sheet5").Select
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("A#")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("Customer")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("Invoice Date")
        .Orientation = xlRowField
        .Position = 3
    End With
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("Invoice  #")
        .Orientation = xlRowField
        .Position = 4
    End With
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("L1 Error")
        .PivotItems("Availability Error").Visible = False
        .PivotItems("Order Entry Error").Visible = False
        .PivotItems("(blank)").Visible = False
    End With
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("L1 Error")
        .Orientation = xlRowField
        .Position = 5
    End With
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("L3 Error")
        .Orientation = xlRowField
        .Position = 6
    End With
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("Responsible")
        .Orientation = xlRowField
        .Position = 7
    End With
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("L1 Error")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable2").AddDataField ActiveSheet.PivotTables( _
        "PivotTable2").PivotFields("L3 Error"), "Count of L3 Error", xlCount
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("L3 Error")
        .Orientation = xlRowField
        .Position = 6
    End With
    Sheets("Sheet5").Name = "Execution Error Pivot"
    ActiveSheet.PivotTables("PivotTable2").Name = "ExecutionError"
    With ActiveSheet.PivotTables("ExecutionError")
        .InGridDropZones = True
        .RowAxisLayout xlTabularRow
    End With
    ActiveSheet.PivotTables("ExecutionError").PivotFields("Customer").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("ExecutionError").PivotFields("Invoice Date"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("ExecutionError").PivotFields("Invoice  #").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("ExecutionError").PivotFields("Responsible").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    Cells.Select
    Cells.EntireColumn.AutoFit
End Sub


-----------------------------------------------------------------------------------------------
Sub AvailabilityErrors()
'
' AvailabilityErrors Macro
' Creates the Availability Errors Tab and Pivot Table
'

'
    Sheets.Add
    ActiveWorkbook.Worksheets("ALLERRORS").PivotTables("PlatinumPivot"). _
        PivotCache.CreatePivotTable TableDestination:="Sheet6!R3C1", TableName:= _
        "PivotTable3", DefaultVersion:=xlPivotTableVersion15
    Sheets("Sheet6").Select
    Cells(3, 1).Select
    Sheets("Sheet6").Select
    Sheets("Sheet6").Name = "Availability Errors"
    With ActiveSheet.PivotTables("PivotTable3").PivotFields("A#")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable3").PivotFields("Customer")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("PivotTable3").PivotFields("Invoice Date")
        .Orientation = xlRowField
        .Position = 3
    End With
    With ActiveSheet.PivotTables("PivotTable3").PivotFields("Invoice  #")
        .Orientation = xlRowField
        .Position = 4
    End With
    With ActiveSheet.PivotTables("PivotTable3").PivotFields("L1 Error")
        .Orientation = xlRowField
        .Position = 5
    End With
    With ActiveSheet.PivotTables("PivotTable3").PivotFields("L1 Error")
        .PivotItems("Execution Error").Visible = False
        .PivotItems("Order Entry Error").Visible = False
        .PivotItems("(blank)").Visible = False
    End With
    With ActiveSheet.PivotTables("PivotTable3").PivotFields("L3 Error")
        .Orientation = xlRowField
        .Position = 6
    End With
    With ActiveSheet.PivotTables("PivotTable3").PivotFields("Responsible")
        .Orientation = xlRowField
        .Position = 7
    End With
    With ActiveSheet.PivotTables("PivotTable3").PivotFields("L1 Error")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable3").AddDataField ActiveSheet.PivotTables( _
        "PivotTable3").PivotFields("L3 Error"), "Count of L3 Error", xlCount
    With ActiveSheet.PivotTables("PivotTable3").PivotFields("Responsible")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable3").PivotFields("Customer")
        .Orientation = xlRowField
        .Position = 2
    End With
    ActiveSheet.PivotTables("PivotTable3").Name = "AvailabilityErrors"
    With ActiveSheet.PivotTables("AvailabilityErrors")
        .InGridDropZones = True
        .RowAxisLayout xlTabularRow
    End With
    ActiveSheet.PivotTables("AvailabilityErrors").PivotFields("Customer"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("AvailabilityErrors").PivotFields("A#").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("AvailabilityErrors").PivotFields("Invoice Date"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("AvailabilityErrors").PivotFields("Invoice  #"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    Cells.Select
    Cells.EntireColumn.AutoFit
    ActiveSheet.PivotTables("AvailabilityErrors").PivotFields("Customer"). _
        Subtotals = Array(True, False, False, False, False, False, False, False, False, False, _
        False, False)
    Cells.EntireColumn.AutoFit
End Sub

-----------------------------------------------------------------------------------------------
Sub Editor()
'
' Editor Macro
' Use this to record macros and refine code
'

'
    With ActiveSheet.PivotTables("ALLERRORS").PivotFields("A#")
        .PivotItems("").Visible = False
    End With
End Sub


-----------------------------------------------------------------------------------------------
Sub BuyerTabsSFD()
'
' BuyerTabsSFD Macro
' Copy and Paste the WMSLOT Data to be separated by buyer
'

'
    Sheets("Sheet2").Select
    Sheets("Sheet2").Name = "ALL SFD"
    Columns("A:A").Select
    Range("A2").Activate
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Range("A1").Select
    ActiveSheet.Range("$A$1:$AS$1053").AutoFilter Field:=1, Criteria1:= _
        "Danae Humbard"
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet1").Select
    Sheets("Sheet1").Name = "DANAE"
    Sheets("ALL SFD").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("DANAE").Select
    ActiveSheet.Paste
    Cells.Select
    Cells.EntireColumn.AutoFit
    Application.CutCopyMode = False
    Selection.AutoFilter
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    Sheets("ALL SFD").Select
    Range("A1").Select
    ActiveSheet.Range("$A$1:$AS$1053").AutoFilter Field:=1, Criteria1:= _
        "DISCONTINUED"
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet2").Select
    Sheets("Sheet2").Name = "MISC."
    ActiveSheet.Paste
    Cells.Select
    Application.CutCopyMode = False
    Selection.AutoFilter
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    Sheets("ALL SFD").Select
    Range("A1").Select
    ActiveSheet.Range("$A$1:$AS$1053").AutoFilter Field:=1, Criteria1:= _
        "Penny Reno"
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet3").Select
    Sheets("Sheet3").Name = "PENNY"
    ActiveSheet.Paste
    Cells.Select
    Application.CutCopyMode = False
    Selection.AutoFilter
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    Sheets("ALL SFD").Select
    Range("A1").Select
    ActiveSheet.Range("$A$1:$AS$1053").AutoFilter Field:=1, Criteria1:= _
        "Todd Kamp"
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet4").Select
    Sheets("Sheet4").Name = "TODD"
    ActiveSheet.Paste
    Cells.Select
    Application.CutCopyMode = False
    Selection.AutoFilter
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    Sheets("ALL SFD").Select
    Range("A1").Select
    ActiveSheet.Range("$A$1:$AS$1053").AutoFilter Field:=1, Criteria1:= _
        "Vincent Romi"
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet5").Select
    Sheets("Sheet5").Name = "VINCENT"
    ActiveSheet.Paste
    Cells.Select
    Application.CutCopyMode = False
    Selection.AutoFilter
    Selection.ColumnWidth = 8.57
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    Sheets("ALL SFD").Select
    Range("A1").Select
    ActiveSheet.Range("$A$1:$AS$1053").AutoFilter Field:=1
    Sheets("ALL SFD").Select
    Range("A1").Select
    ActiveWorkbook.SaveCopyAs ("C:\Users\PRPowell\Desktop\NonShared\WMS Aging Files\Springfield\" & Format("120 Day Aging - Springfield mm.dd.yyyy") & ".xlsm")
End Sub