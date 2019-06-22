let
    Source = Excel.CurrentWorkbook(){[Name="config"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"parameter", type text}, {"value", type text}})
in
    #"Changed Type"

let
    Source = Folder.Files("C:\Users\prp12.000\GitHub-Repos\Excel"),
    #"Removed Columns" = Table.RemoveColumns(Source,{"Content"}),
    #"Filtered Rows" = Table.SelectRows(#"Removed Columns", each ([Extension] <> "" and [Extension] <> ".idx" and [Extension] <> ".pack" and [Extension] <> ".sample")),
    #"Replaced Value" = Table.ReplaceValue(#"Filtered Rows","C:\Users\prp12.000\GitHub-Repos\Excel","https://github.com/prp1277/Excel/tree/master",Replacer.ReplaceText,{"Folder Path"}),
    #"Replaced Value1" = Table.ReplaceValue(#"Replaced Value","\","/",Replacer.ReplaceText,{"Folder Path"}),
    #"Merged Columns" = Table.CombineColumns(#"Replaced Value1",{"Folder Path", "Name"},Combiner.CombineTextByDelimiter("", QuoteStyle.None),"Index"),
    #"Removed Other Columns" = Table.SelectColumns(#"Merged Columns",{"Index"})
in
    #"Removed Other Columns"
