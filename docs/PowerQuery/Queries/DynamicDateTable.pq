let
    Source = #date(2019,1,1),
    Custom1 = List.Dates(Source, Number.From(DateTime.LocalNow())- Number.From(Source) ,#duration(1,0,0,0)),
    #"Converted to Table" = Table.FromList(Custom1, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    #"Duplicated Column" = Table.DuplicateColumn(#"Converted to Table", "Column1", "Column1 - Copy"),
    #"Duplicated Column1" = Table.DuplicateColumn(#"Duplicated Column", "Column1", "Column1 - Copy.1"),
    #"Duplicated Column2" = Table.DuplicateColumn(#"Duplicated Column1", "Column1", "Column1 - Copy.2"),
    #"Duplicated Column3" = Table.DuplicateColumn(#"Duplicated Column2", "Column1", "Column1 - Copy.3"),
    #"Duplicated Column4" = Table.DuplicateColumn(#"Duplicated Column3", "Column1", "Column1 - Copy.4"),
    #"Duplicated Column5" = Table.DuplicateColumn(#"Duplicated Column4", "Column1", "Column1 - Copy.5"),
    #"Extracted Year" = Table.TransformColumns(#"Duplicated Column5",{{"Column1 - Copy", Date.Year, Int64.Type}}),
    #"Extracted Month" = Table.TransformColumns(#"Extracted Year",{{"Column1 - Copy.1", Date.Month, Int64.Type}}),
    #"Calculated Quarter" = Table.TransformColumns(#"Extracted Month",{{"Column1 - Copy.2", Date.QuarterOfYear, Int64.Type}}),
    #"Calculated Week of Year" = Table.TransformColumns(#"Calculated Quarter",{{"Column1 - Copy.3", Date.WeekOfYear, Int64.Type}}),
    #"Extracted Day" = Table.TransformColumns(#"Calculated Week of Year",{{"Column1 - Copy.4", Date.Day, Int64.Type}}),
    #"Extracted Day Name" = Table.TransformColumns(#"Extracted Day", {{"Column1 - Copy.5", each Date.DayOfWeekName(_), type text}}),
    #"Renamed Columns" = Table.RenameColumns(#"Extracted Day Name",{{"Column1 - Copy", "Year"}, {"Column1 - Copy.1", "Month"}, {"Column1 - Copy.2", "Quarter"}, {"Column1 - Copy.3", "Week"}, {"Column1 - Copy.4", "Day"}, {"Column1", "Date"}, {"Column1 - Copy.5", "Name of Day"}}),
    #"Changed Type" = Table.TransformColumnTypes(#"Renamed Columns",{{"Date", type date}})
in
    #"Changed Type"
