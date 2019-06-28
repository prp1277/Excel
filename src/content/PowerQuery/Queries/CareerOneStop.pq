/*
* Parameters
*/

/* baseURL */
"https://api.careeronestop.org/v1/jobsearch/ijJDYCadAcEJZ5e/" meta [IsParameterQuery=true, Type="Any", IsParameterQueryRequired=true]

/* JobTitle */
"Business%20Intelligence" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]

/* QS */
"/68144/10/company/ASC/1/200/30?source=NLx&showFilters=true" meta [IsParameterQuery=true, Type="Any", IsParameterQueryRequired=true]

/* forEach */
let
  Source = (query as text) =>
    let
      Source = Json.Document(Web.Contents(baseURL & query & QS, [Headers=[Authorization="Bearer //Your Auth"]])),
      #"Converted to Table" = Record.ToTable(Source)
    in
      #"Converted to Table"
in
  Source

/* financialAnalyst */
let
    Source = forEach("financial%20analyst"),
    #"Filtered Rows" = Table.SelectRows(Source, each ([Name] = "Jobs")),
    #"Expanded Value" = Table.ExpandListColumn(#"Filtered Rows", "Value"),
    #"Expanded Value1" = Table.ExpandRecordColumn(#"Expanded Value", "Value", {"JvId", "JobTitle", "Company", "AccquisitionDate", "URL", "Location", "Fc"}, {"JvId", "JobTitle", "Company", "AccquisitionDate", "URL", "Location", "Fc"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded Value1",{{"JvId", type text}, {"JobTitle", type text}, {"Company", type text}, {"AccquisitionDate", type datetimezone}, {"URL", type text}, {"Location", type text}, {"Fc", type logical}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"Name", "JvId", "Location", "Fc"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Removed Columns",{"AccquisitionDate", "Company", "JobTitle", "URL"}),
    #"Added Prefix" = Table.TransformColumns(#"Reordered Columns", {{"URL", each "=HYPERLINK("" & _, type text}}),
    #"Added Suffix" = Table.TransformColumns(#"Added Prefix", {{"URL", each _ & "", "Apply Now")", type text}})
in
    #"Added Suffix"

/* businessAnalyst */
let
    Source = forEach("business%20analyst"),
    #"Filtered Rows" = Table.SelectRows(Source, each ([Name] = "Jobs")),
    #"Expanded Value" = Table.ExpandListColumn(#"Filtered Rows", "Value"),
    #"Expanded Value1" = Table.ExpandRecordColumn(#"Expanded Value", "Value", {"JvId", "JobTitle", "Company", "AccquisitionDate", "URL", "Location", "Fc"}, {"JvId", "JobTitle", "Company", "AccquisitionDate", "URL", "Location", "Fc"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded Value1",{{"JvId", type text}, {"JobTitle", type text}, {"Company", type text}, {"AccquisitionDate", type datetimezone}, {"URL", type text}, {"Location", type text}, {"Fc", type logical}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"Name", "JvId"}),
    #"Duplicated Column" = Table.DuplicateColumn(#"Removed Columns", "URL", "URL - Copy"),
    #"Added Prefix" = Table.TransformColumns(#"Duplicated Column", {{"URL - Copy", each "=HYPERLINK("" & _, type text}}),
    #"Added Suffix" = Table.TransformColumns(#"Added Prefix", {{"URL - Copy", each _ & "", "Apply Now")", type text}}),
    #"Renamed Columns" = Table.RenameColumns(#"Added Suffix",{{"URL - Copy", "Apply Now"}}),
    #"Removed Columns1" = Table.RemoveColumns(#"Renamed Columns",{"URL", "Location", "Fc"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Removed Columns1",{"AccquisitionDate", "Company", "JobTitle", "Apply Now"})
in
    #"Reordered Columns"
