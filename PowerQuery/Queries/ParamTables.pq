let
    Source = Excel.CurrentWorkbook(){[Name="MyTable"]}[Content],
    #"Added Custom" = Table.AddColumn(Source, "File Path", each fnGetParameter("File Path")),
    #"Added Custom1" = Table.AddColumn(#"Added Custom", "First Date", each fnGetParameter("Start Date")),
    #"Changed Type" = Table.TransformColumnTypes(#"Added Custom1",{{"File Path", type text}, {"First Date", type date}, {"MyTable", type text}})
in
    #"Changed Type"
/** fnGetParameter */
(ParameterName as text) =>
let
    ParamSource = Excel.CurrentWorkbook(){[Name="Parameters"]}[Content],
    ParamRow = Table.SelectRows(ParamSource, each([Parameter] = ParameterName)),
    Value = if Table.IsEmpty(ParamRow)=true
    then null
    else Record.Field(ParamRow{0}, "Value")
in
Value