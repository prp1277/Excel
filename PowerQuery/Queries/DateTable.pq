let
    Source = {DateTime.LocalNow() - #datetime(1970,01,01,0,0,0)},
    #"Converted to Table" = Table.FromList(Source, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    #"Duplicated Column" = Table.DuplicateColumn(#"Converted to Table", "Column1", "Column1 - Copy"),
    #"Renamed Columns" = Table.RenameColumns(#"Duplicated Column",{{"Column1", "Unix Time"}, {"Column1 - Copy", "Duration"}}),
    #"Calculated Total Seconds" = Table.TransformColumns(#"Renamed Columns",{{"Unix Time", Duration.TotalSeconds, type text}}),
    #"Changed Type" = Table.TransformColumnTypes(#"Calculated Total Seconds",{{"Duration", type duration}, {"Unix Time", type number}}),
    #"Added Custom" = Table.AddColumn(#"Changed Type", "Now", each DateTime.LocalNow()),
    #"Changed Type1" = Table.TransformColumnTypes(#"Added Custom",{{"Now", type datetimezone}})
in
    #"Changed Type1"