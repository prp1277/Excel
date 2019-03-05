//-----------------------------------------------------------------------------
// ParamStrings
//-----------------------------------------------------------------------------
// baseURL
"https://swapi.co/api/" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]

// query
"planets" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]

// page - Adds Pagination
"/6" meta [IsParameterQuery=true, Type="Any", IsParameterQueryRequired=true]

//-----------------------------------------------------------------------------
// Validation
//-----------------------------------------------------------------------------
// results as a string
let
    Source = baseURL & query & page
in
    Source

//-----------------------------------------------------------------------------
// Results
//-----------------------------------------------------------------------------
// Results as list
let
    Source = Json.Document(Web.Contents(baseURL & query)),
    results = Source[results]
in
    results

//-----------------------------------------------------------------------------
// Functions
//-----------------------------------------------------------------------------
// forEach(`${query}`)
let
    Source = (query as text) => let
        Source = Json.Document(Web.Contents(baseURL & query)),
        results = Source[results]
    in
        results
in
    Source

//-----------------------------------------------------------------------------
// Functions
//-----------------------------------------------------------------------------
// Paginate
let
    Source = Json.Document(Web.Contents(baseURL & query & page))
in
    Source

//-----------------------------------------------------------------------------
// Invoked Functions
//-----------------------------------------------------------------------------
// "planets" 
let
    Source = forEach("planets")
in
    Source
    
//-----------------------------------------------------------------------------
// Invoked Functions
//-----------------------------------------------------------------------------
// "people"
let
    Source = forEach("people")
in
    Source