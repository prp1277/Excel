Sub ParameterTable()
' ParameterTable Macro

Range("A1").Value = "parameter"
Range("B1").Value = "value"
Range("A2").Value = "thisNotebook"
Range("B2").Value = "=CELL(""filename"")"

Range("A1").Activate

    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    ActiveSheet.ListObjects.Add(xlSrcRange, Range(Selection), , xlYes).Name = _
        "config"
    ChDir "C:\Users\prp12.000\OneDrive\Apps\Excel"
    ActiveWorkbook.SaveAs Filename:= _
        "https://d.docs.live.net/b27236921334e482/Apps/Excel/pTable.xlsm", FileFormat _
        :=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
End Sub
'    ActiveWorkbook.Queries.Add Name:="config", Formula:= _
'        "let" & Chr(13) & "" & Chr(10) & "    Source = Excel.CurrentWorkbook(){[Name=""config""]}[Content]," & Chr(13) & "" & Chr(10) & "    #""Changed Type"" = Table.TransformColumnTypes(Source,{{""parameter"", type text}, {""value"", type text}})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    #""Changed Type"""
'    Workbooks("Book1").Connections.Add2 "Query - config", _
'        "Connection to the 'config' query in the workbook.", _
'        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=config;Extended Properties=" _
'        , """config""", 6, True, False
'    Range("config[[#Headers],[parameter]]").Select
