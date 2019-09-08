/** Source = Folder */
let
    Source = Folder.Files("C:\Users\prp12.000\prp1277-Github\prp1277.github.io\src\pages\docs\excel")
in
    Source

/** Transformation */
let
    Source = Folder.Files("C:\Users\prp12.000\prp1277-Github\prp1277.github.io\src\pages\docs\excel"),
    #"C:\Users\prp12 000\prp1277-Github\prp1277 github io\src\pages\docs\excel\_gh-flavored-markdown md" = Source{[#"Folder Path"="C:\Users\prp12.000\prp1277-Github\prp1277.github.io\src\pages\docs\excel\",Name="gh-flavored-markdown.md"]}[Content],
    #"Imported CSV" = Csv.Document(#"C:\Users\prp12 000\prp1277-Github\prp1277 github io\src\pages\docs\excel\_gh-flavored-markdown md",[Delimiter="#(tab)", Encoding=1252, QuoteStyle=QuoteStyle.Csv]),
    #"Kept Range of Rows" = Table.Range(#"Imported CSV",1,3),
    #"Split Column by Delimiter" = Table.SplitColumn(#"Kept Range of Rows", "Column1", Splitter.SplitTextByDelimiter(":", QuoteStyle.None), {"Column1.1", "Column1.2"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Split Column by Delimiter",{{"Column1.1", type text}, {"Column1.2", type text}}),
    #"Replaced Value" = Table.ReplaceValue(#"Changed Type","[","",Replacer.ReplaceText,{"Column1.2"}),
    #"Replaced Value1" = Table.ReplaceValue(#"Replaced Value","]","",Replacer.ReplaceText,{"Column1.2"}),
    #"Transposed Table" = Table.Transpose(#"Replaced Value1"),
    #"Promoted Headers" = Table.PromoteHeaders(#"Transposed Table", [PromoteAllScalars=true]),
    #"Changed Type1" = Table.TransformColumnTypes(#"Promoted Headers",{{"title", type text}, {"date", type text}, {"tags", type text}}),
    #"Split Column by Delimiter1" = Table.SplitColumn(#"Changed Type1", "tags", Splitter.SplitTextByDelimiter(",", QuoteStyle.Csv), {"tags.1", "tags.2", "tags.3"}),
    #"Changed Type2" = Table.TransformColumnTypes(#"Split Column by Delimiter1",{{"tags.1", type text}, {"tags.2", type text}, {"tags.3", type text}}),
    #"Trimmed Text" = Table.TransformColumns(#"Changed Type2",{{"title", Text.Trim, type text}, {"tags.3", Text.Trim, type text}, {"tags.2", Text.Trim, type text}, {"tags.1", Text.Trim, type text}, {"date", Text.Trim, type text}})
in
    #"Trimmed Text"