/** ------------------------------------------------------------------------- *
* Get Data
* -------------------------------------------------------------------------- */

let
Parameter = Excel.CurrentWorkbook(){[Name="Parameters"]}[Content],
URL = Parameter{0}[Value],
Source = Json.Document(Web.Contents(URL)),
    #"Converted to Table" = Record.ToTable(Source)
in
    #"Converted to Table"

/** ------------------------------------------------------------------------- *
* News
* -------------------------------------------------------------------------- */
let
    Source = URLQuery,
    #"Expanded Value" = Table.ExpandRecordColumn(
        Source, "Value", 
        {"news"}, 
        {"news"}
    ),
    #"Expanded news" = Table.ExpandListColumn(
        #"Expanded Value", "news"
    ),
    #"Expanded news1" = Table.ExpandRecordColumn(
        #"Expanded news", "news", {
        "datetime", "headline", "source", "url", "summary", "related", "image"
        }, {
            "datetime", "headline", "source", "url", "summary", "related", "image"
    }),
    #"Changed Type" = Table.TransformColumnTypes(
        #"Expanded news1",
    {
        {"datetime", type datetimezone}, 
        {"headline", type text}, 
        {"source", type text}, 
        {"url", type text}, 
        {"summary", type text}, 
        {"related", type text}, 
        {"image", type text}}
    ),
    #"Trimmed Text" = Table.TransformColumns(
        #"Changed Type",
    {
        {"image", Text.Trim, type text},
        {"related", Text.Trim, type text},
        {"summary", Text.Trim, type text},
        {"url", Text.Trim, type text},
        {"source", Text.Trim, type text},
        {"headline", Text.Trim, type text}
    }
),
    #"Cleaned Text" = Table.TransformColumns(#"Trimmed Text",
    {
        {"image", Text.Clean, type text},
        {"related", Text.Clean, type text},
        {"summary", Text.Clean, type text},
        {"url", Text.Clean, type text},
        {"source", Text.Clean, type text},
        {"headline", Text.Clean, type text}
    }
)
in
    #"Cleaned Text"

/** ------------------------------------------------------------------------- * 
* Chart
* -------------------------------------------------------------------------- */
let
    Source = URLQuery,
    #"Expanded Value" = Table.ExpandRecordColumn(
        Source, "Value", {"chart"}, {"chart"}
    ),
    #"Expanded chart" = Table.ExpandListColumn(
        #"Expanded Value", "chart"
        ),
    #"Expanded chart1" = Table.ExpandRecordColumn(
        #"Expanded chart", "chart",
        {"date", "open", "high", "low", "close", "volume", "unadjustedVolume", "change", "changePercent", "label", "changeOverTime"},
        {"date", "open", "high", "low", "close", "volume", "unadjustedVolume", "change", "changePercent", "label", "changeOverTime"}
        ),
    #"Changed Type" = Table.TransformColumnTypes(
        #"Expanded chart1",{
            {"date", type date},
            {"open", Currency.Type},
            {"high", Currency.Type},
            {"low", Currency.Type},
            {"close", Currency.Type},
            {"volume", Int64.Type},
            {"unadjustedVolume", Int64.Type},
            {"label", type date},
            {"change", Currency.Type},
            {"changePercent", Percentage.Type},
            {"changeOverTime", Percentage.Type}
    }),
    #"Divided Column" = Table.TransformColumns(
        #"Changed Type", {
            {"changePercent", each _ / 100, Percentage.Type}
            }
        ),
    #"Sorted Rows" = Table.Sort(#"Divided Column",{
        {"Name", Order.Ascending},
        {"date", Order.Ascending}
    })
in
    #"Sorted Rows"


/** ------------------------------------------------------------------------- *
* 2018-10-18 - IEX Query 
******************************************************************************/

/** ------------------------------------------------------------------------- *
* baseURL
* -------------------------------------------------------------------------- */

let
    Source = Table.FromRows(
        Json.Document(
            Binary.Decompress(
                Binary.FromText(
                    "i45WSkosTg0N8rFS0lHKKCkpKLbS108syNTLTK0oKUpMycxL10vOz9U31DNQio0FAA==",
                    BinaryEncoding.Base64
                ),
                Compression.Deflate
            )
        ), 
        let _t = ((type text) meta [Serialized.Text = true])
        in type table [key = _t, value = _t]
        ),
        #"Changed Type" = Table.TransformColumnTypes(Source,{{"key", type text}, {"value", type text}}),
        value = #"Changed Type"[value]
in
    value

/** ------------------------------------------------------------------------- *
* url ()
* -------------------------------------------------------------------------- */

"https://api.iextrading.com/1.0" meta [
    IsParameterQuery=true, 
    ExpressionIdentifier=baseURL, 
    Type="Text", 
    IsParameterQueryRequired=true
    ]

/** ------------------------------------------------------------------------- *
* params ()
* -------------------------------------------------------------------------- */

"/stock/market/" meta [
    IsParameterQuery=true,
    Type="Text",
    IsParameterQueryRequired=true
    ]

/** ------------------------------------------------------------------------- *
* batch ()
* -------------------------------------------------------------------------- */

"batch?" meta [
    IsParameterQuery=true,
    Type="Text",
    IsParameterQueryRequired=true
    ]

/** ------------------------------------------------------------------------- *
* collection ()
* -------------------------------------------------------------------------- */

"/collection/sector?collectionName=Health%20Care" meta [
    IsParameterQuery=true,
    Type="Any",
    IsParameterQueryRequired=true
    ]

/** ------------------------------------------------------------------------- *
* symbols ()
* -------------------------------------------------------------------------- */

"&symbols=ndaq,mdb,cag,pfe,msft" meta [
    IsParameterQuery=true,
    Type="Text",
    IsParameterQueryRequired=true
    ]

/** ------------------------------------------------------------------------- *
* types ()
* -------------------------------------------------------------------------- */

"&types=news,quote,chart,company,logo" meta [
    IsParameterQuery=true,
    Type="Any",
    IsParameterQueryRequired=true
    ]

/** ------------------------------------------------------------------------- *
* range()
* -------------------------------------------------------------------------- */

"&range=1y" meta [
    IsParameterQuery=true,
    Type="Any",
    IsParameterQueryRequired=true
    ]

/** ------------------------------------------------------------------------- *
* last ()
* -------------------------------------------------------------------------- */

"&last=10" meta [
    IsParameterQuery=true,
    Type="Any",
    IsParameterQueryRequired=true
    ]


/******************************************************************************
*******************************************************************************
* Requests
* -----------------------------------------------------------------------------
* Request
* -------------------------------------------------------------------------- */

let
    Source = Json.Document(
        Web.Contents(
            url & params & batch & symbols & types & range & last
        )
    ),
    #"Converted to Table" = Record.ToTable(Source)
in
    #"Converted to Table"

/** ------------------------------------------------------------------------- *
* collections
* -------------------------------------------------------------------------- */

let
    Source = Json.Document(
        Web.Contents(
            url & params & collection
        )
    ),
    #"Converted to Table" = Table.FromList(
        Source, Splitter.SplitByNothing(), 
        null, 
        null, 
        ExtraValues.Error
    )
in
    #"Converted to Table"

/******************************************************************************
* Other Queries
******************************************************************************/

/** ------------------------------------------------------------------------- *
* news
* -------------------------------------------------------------------------- */

let
    Source = Request,
    #"Expanded Value" = Table.ExpandRecordColumn(
        Source, "Value",
        {"news"},
        {"news"}
    ),
    #"Expanded news" = Table.ExpandListColumn
    (
        #"Expanded Value", "news"
    ),
    #"Expanded news1" = Table.ExpandRecordColumn(
        #"Expanded news", "news",
        {"datetime", "headline", "source", "url", "summary", "related", "image"},
        {"datetime", "headline", "source", "url", "summary", "related", "image"}
    ),
    #"Trimmed Text" = Table.TransformColumns(
        #"Expanded news1",{
            {"datetime", Text.Trim, type text}, 
            {"headline", Text.Trim, type text}, 
            {"source", Text.Trim, type text}, 
            {"url", Text.Trim, type text}, 
            {"summary", Text.Trim, type text}, 
            {"related", Text.Trim, type text}, 
            {"image", Text.Trim, type text}
        }
    ),
    #"Cleaned Text" = Table.TransformColumns(
        #"Trimmed Text",{
            {"datetime", Text.Clean, type text}, 
            {"headline", Text.Clean, type text}
            {"source", Text.Clean, type text}, 
            {"url", Text.Clean, type text}, 
            {"summary", Text.Clean, type text}, 
            {"related", Text.Clean, type text}, 
            {"image", Text.Clean, type text}
        }
    ),
    #"Changed Type" = Table.TransformColumnTypes(
        #"Cleaned Text", {
            {"datetime", type datetimezone}
        }
    )
in
    #"Changed Type"

/** ------------------------------------------------------------------------- *
* quote
* -------------------------------------------------------------------------- */

let
    Source = Request,
    #"Expanded Value" = Table.ExpandRecordColumn(
        Source, "Value", {"chart"}, {"chart"}
    ),
    #"Expanded chart" = Table.ExpandListColumn(
        #"Expanded Value", "chart"
    ),
    #"Expanded chart1" = Table.ExpandRecordColumn(
        #"Expanded chart", "chart", {
            "date", "open", "high", "low", "close", "change", "changePercent", "label", "changeOverTime"}, {"date", "open", "high", "low", "close", "change", "changePercent", "label", "changeOverTime"
        }
    ),
    #"Changed Type" = Table.TransformColumnTypes(
        #"Expanded chart1",{
            {"open", Currency.Type},
            {"high", Currency.Type}, 
            {"low", Currency.Type}, 
            {"close", Currency.Type}, 
            {"change", Currency.Type}, 
            {"changeOverTime", Percentage.Type}, 
            {"changePercent", Percentage.Type}, 
            {"label", type date}
        }
    ),
    #"Divided Column" = Table.TransformColumns(
        #"Changed Type", {
            {"changePercent", each _ / 100, Percentage.Type}
        }
    ),
    #"Changed Type1" = Table.TransformColumnTypes(
        #"Divided Column",{
            {"date", type date}
        }
    )
in
    #"Changed Type1"

/** ------------------------------------------------------------------------- *
* chart
* -------------------------------------------------------------------------- */

let
    Source = Request,
    #"Expanded Value" = Table.ExpandRecordColumn(
        Source, "Value", {"quote"}, {"quote"}
    ),
    #"Expanded quote" = Table.ExpandRecordColumn(
        #"Expanded Value", "quote", {
            "calculationPrice", "open", "openTime", "close", "closeTime", "high", "low", "latestPrice", "latestSource", "latestTime", "latestUpdate", "latestVolume", "delayedPrice", "delayedPriceTime", "extendedPrice", "previousClose", "change", "changePercent", "avgTotalVolume", "marketCap", "peRatio", "week52High", "week52Low", "ytdChange"}, {"calculationPrice", "open", "openTime", "close", "closeTime", "high", "low", "latestPrice", "latestSource", "latestTime", "latestUpdate", "latestVolume", "delayedPrice", "delayedPriceTime", "extendedPrice", "previousClose", "change", "changePercent", "avgTotalVolume", "marketCap", "peRatio", "week52High", "week52Low", "ytdChange"
        }
    ),
    #"Changed Type" = Table.TransformColumnTypes(
        #"Expanded quote",{
            {"calculationPrice", type text},
            {"open", Currency.Type},
            {"close", Currency.Type},
            {"high", Currency.Type},
            {"low", Currency.Type},
            {"latestPrice", Currency.Type},
            {"latestSource", type text},
            {"latestTime", type datetime},
            {"latestVolume", Int64.Type},
            {"delayedPrice", Currency.Type},
            {"extendedPrice", Currency.Type},
            {"previousClose", Currency.Type},
            {"change", Currency.Type},
            {"changePercent", Percentage.Type},
            {"avgTotalVolume", Int64.Type},
            {"marketCap", Currency.Type},
            {"peRatio", type number},
            {"week52High", Currency.Type},
            {"week52Low", Currency.Type},
            {"ytdChange", Percentage.Type},
            {"openTime", type number},
            {"closeTime", type number}, 
            {"latestUpdate", type number}, 
            {"delayedPriceTime", type number}
        }
    )
in
    #"Changed Type"

/** ------------------------------------------------------------------------- *
* range
* -------------------------------------------------------------------------- */

let
    Source = Request,
    #"Expanded Value" = Table.ExpandRecordColumn(
        Source, "Value", {"company"}, {"company"}
    ),
    #"Expanded company" = Table.ExpandRecordColumn(
        #"Expanded Value", "company", {"companyName", "exchange", "industry", "website", "description", "CEO", "issueType", "sector", "tags"}, {"companyName", "exchange", "industry", "website", "description", "CEO", "issueType", "sector", "tags"}
    ),
    #"Extracted Values" = Table.TransformColumns(
        #"Expanded company", {"tags", each Text.Combine(
            List.Transform(_, Text.From), ","
        ), type text}
    ),
    #"Changed Type" = Table.TransformColumnTypes(
        #"Extracted Values",{
            {"sector", type text}, 
            {"issueType", type text}, 
            {"CEO", type text}, 
            {"description", type text}, 
            {"website", type text}, 
            {"industry", type text}, 
            {"exchange", type text}, 
            {"companyName", type text}
        }
    )
in
    #"Changed Type"

/** ------------------------------------------------------------------------- *
* type
* -------------------------------------------------------------------------- */

let
    Source = Request,
    #"Expanded Value" = Table.ExpandRecordColumn(
        Source, "Value", {"logo"}, {"logo"}
    ),
    #"Expanded logo" = Table.ExpandRecordColumn(
        #"Expanded Value", "logo", {"url"}, {"url"}
    ),
    #"Changed Type" = Table.TransformColumnTypes(
        #"Expanded logo",{{"url", type text}}
    )
in
    #"Changed Type"

/** ------------------------------------------------------------------------- *
* last
* -------------------------------------------------------------------------- */

let
    Source = DateTime.LocalNow(),
    #"Converted to Table" = #table(1, {{Source}}),
    #"Renamed Columns" = Table.RenameColumns(
        #"Converted to Table",{
            {"Column1", "DateTime LocalNow"}
        }
    ),
    #"Changed Type" = Table.TransformColumnTypes(
        #"Renamed Columns",{
            {"DateTime LocalNow", type datetime}
        }
    ),
    #"Added Custom" = Table.AddColumn(
        #"Changed Type", "DateFrom POSIX", each [DateTime LocalNow] - #datetime(1970,01,01,00,00,00)),
    #"Changed Type1" = Table.TransformColumnTypes(
        #"Added Custom",{
            {"DateFrom POSIX", type duration}
        }
    ),
    #"Inserted Total Seconds" = Table.AddColumn(
        #"Changed Type1", "NOW() Hash", each Duration.TotalSeconds(
            [DateFrom POSIX]
        ), type number),
    #"Inserted Date" = Table.AddColumn(
        #"Inserted Total Seconds", "Date", each DateTime.Date(
            [DateTime LocalNow]
        ), type date),
    #"Inserted Time" = Table.AddColumn(
        #"Inserted Date", "Time", each DateTime.Time(
            [DateTime LocalNow]
        ), type time),
    #"Changed Type2" = Table.TransformColumnTypes(
        #"Inserted Time",{
            {"Time", type time}
        }
    )
in
    #"Changed Type2"

/** ------------------------------------------------------------------------- *
* company
* -------------------------------------------------------------------------- */

let
    Source = collections,
    #"Expanded Column1" = Table.ExpandRecordColumn(
        Source, "Column1", {
            "symbol", "companyName", "primaryExchange", "sector", "calculationPrice", "open", "openTime", "close", "closeTime", "high", "low", "latestPrice", "latestSource", "latestTime", "latestUpdate", "latestVolume", "iexRealtimePrice", "iexRealtimeSize", "iexLastUpdated", "delayedPrice", "delayedPriceTime", "extendedPrice", "extendedChange", "extendedChangePercent", "extendedPriceTime", "previousClose", "change", "changePercent", "iexMarketPercent", "iexVolume", "avgTotalVolume", "iexBidPrice", "iexBidSize", "iexAskPrice", "iexAskSize", "marketCap", "peRatio", "week52High", "week52Low", "ytdChange"}, {"symbol", "companyName", "primaryExchange", "sector", "calculationPrice", "open", "openTime", "close", "closeTime", "high", "low", "latestPrice", "latestSource", "latestTime", "latestUpdate", "latestVolume", "iexRealtimePrice", "iexRealtimeSize", "iexLastUpdated", "delayedPrice", "delayedPriceTime", "extendedPrice", "extendedChange", "extendedChangePercent", "extendedPriceTime", "previousClose", "change", "changePercent", "iexMarketPercent", "iexVolume", "avgTotalVolume", "iexBidPrice", "iexBidSize", "iexAskPrice", "iexAskSize", "marketCap", "peRatio", "week52High", "week52Low", "ytdChange"
        }
    )
in
    #"Expanded Column1"