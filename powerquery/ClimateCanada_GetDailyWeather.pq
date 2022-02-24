/**
* Queries climate.weather.gc.ca for daily bulk data (in csv format) for a given climate station and date range. Station
* IDs are available from climate.weather.gc.ca google drive
*   https://drive.google.com/drive/folders/1WJCDEU34c60IfOnG4rv5EPZ4IhhW9vZH/
*
* the csv schema is hard-coded within this query and will need adjusted if the format changes.
*
* @name ClimateCanada.GetDailyWeather
* @categories weather
* @license MIT (c) 2021 W Bohm
* @author https://github.com/AnHerbWorm
* @source https://github.com/AnHerbWorm/DataAnalysis/tree/main/powerquery
* @result table
*/
let

    /*
    CLIMATE FUNCTIONS
        Functions and Types related to climate.weather.gc.ca

    Climate.CSVSchema: Table type of bulk .csv data downloads
    Climate.ApplySchema: Transform table column types to match those of Climate.CSVSchema
    Climate.MakeURL: Construct url to bulk download the entire year's daily data for a given station id
    Climate.ReadURL: Download data and clean up into a table of type Climate.CSVSchema
    */
    Climate.CSVSchema = type table [
        #"Longitude (x)" =  number,
        #"Latitude (y)" =  number,
        #"Station Name" =  text,
        #"Climate ID" = Int64.Type,
        #"Date/Time" =  date,
        #"Year" = Int64.Type,
        #"Month" = Int64.Type,
        #"Day" = Int64.Type,
        #"Data Quality" =  text,
        #"Max Temp (°C)" =  number,
        #"Max Temp Flag" =  text,
        #"Min Temp (°C)" =  number,
        #"Min Temp Flag" =  text,
        #"Mean Temp (°C)" =  number,
        #"Mean Temp Flag" =  text,
        #"Heat Deg Days (°C)" =  number,
        #"Heat Deg Days Flag" =  text,
        #"Cool Deg Days (°C)" =  number,
        #"Cool Deg Days Flag" =  text,
        #"Total Rain (mm)" =  text,
        #"Total Rain Flag" =  text,
        #"Total Snow (cm)" =  text,
        #"Total Snow Flag" =  text,
        #"Total Precip (mm)" =  number,
        #"Total Precip Flag" =  text,
        #"Snow on Grnd (cm)" = Int64.Type,
        #"Snow on Grnd Flag" =  text,
        #"Dir of Max Gust (10s deg)" = Int64.Type,
        #"Dir of Max Gust Flag" =  text,
        #"Spd of Max Gust (km/h)" = Int64.Type,
        #"Spd of Max Gust Flag" =  text
    ],
    Climate.ApplySchema = (table as table) as table => Table.TransformColumnTypes(
        table,
        List.Transform(Type.TableSchema(Climate.CSVSchema)[Name], each {_, Type.TableColumn(Climate.CSVSchema, _)})
    ),  
    Climate.MakeURL = (stationID as number, year as number) as text => 
        Text.Combine(
            {
                "http://climate.weather.gc.ca/climate_data/bulk_data_e.html?format=csv&",
                Text.Format("stationID=#[id]&Year=#[yr]", [id=stationID, yr=year]),
                "&Month=12&Day=31&timeframe=2"
            },
            ""
        ),
    Climate.ReadURL = (url as text) as table => 
        let
            Source = Csv.Document(
                Web.Contents(url),
                [Delimiter=",", Columns=31, Encoding=65001, QuoteStyle=QuoteStyle.None]
            ),
            PromoteHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
            ApplySchema = Climate.ApplySchema(PromoteHeaders)
        in
            ApplySchema,
    /*
    HELPER FUNCTIONS

    MakeDateRange: Return table of inclusive dates between start/end. Single column is named "Date"
    */
    MakeDateRange = (start as date, end as date) as table => #table(
        type table [Date=date],
        List.Generate(
            ()=>0,
            each _ <= Duration.TotalDays(end - start),
            each _ + 1,
            each {Date.AddDays(start, _)}
        )
    ),
    /* MAIN FUNCTION */
    Main.Func = (stationID as number, startDate as date, endDate as date) as table =>
        let
            AllDates = MakeDateRange(startDate, endDate),
            AllYears = List.Distinct(List.Transform(AllDates[Date], Date.Year)),
            
            URLS = List.Transform(AllYears, each Climate.MakeURL(stationID, _)),
            ClimateData = Table.Combine(List.Transform(URLS, Climate.ReadURL)),
            FilterToRange = Table.Join(
                AllDates,
                {"Date"},
                ClimateData,
                {"Date/Time"}
            ),
            RemoveJoinColumn = Table.RemoveColumns(FilterToRange, "Date")
        in
            RemoveJoinColumn,
    Main.Docs = [
        Documentation.Name = "ClimateCanada_GetDailyWeather",
        Documentation.LongDescription = Text.Combine(
            {
                "Queries climate.weather.gc.ca for daily bulk data (in csv format) for a given climate station and",
                "date range.<br><br>",
                
                "<b>Args<br>------</b><br>",
                "<i>stationID</i>: Climate Station ID. A list of all stationIDs can be found at",
                "https://drive.google.com/drive/folders/1WJCDEU34c60IfOnG4rv5EPZ4IhhW9vZH/<br>",
                "<i>startDate</i>: First date to include in the download (inclusive).<br>",
                "<i>endDate</i>: Last date to include in the download (inclusive).<br><br>",

                "<b>NOTE<br>-------</b><br>",
                "Table schema is set within this query and must be adjusted if the bulk data format changes."
            },
            " "
        ),
        Documentation.Examples = {
            [
                Description = "First two days of 2021, Edmonton International Airport (#27793)",
                Code = "GetDailyClimateData(27793, #date(2021, 1, 1), #date(2021, 1, 2))",
                Result = Text.Combine({
                    "Note: Columns truncated for this example#(lf,lf)",
                    "Table.FromRecords({",
                    "#(lf)  [Date/Time = 2021/01/01, Max Temp = -4.0, Min Temp = -13.4],",
                    "#(lf)  [Date/Time = 2021/01/02, Max Temp = 4.1, Min Temp = -12.0]",
                    "#(lf)})"},
                    ""
                )
            ]
        }
    ],
    Main.Type = Type.ForFunction(
        [
            ReturnType = Climate.CSVSchema,
            Parameters = [
                stationID = type number meta [
                    Documentation.FieldCaption = "Climate Station ID",
                    Documentation.SampleValues = {27793}
                ],
                startDate = type date meta [
                    Documentation.FieldCaption = "First Date to Import"
                ],
                endDate = type date meta [
                    Documentation.FieldCaption = "Last Date to Import"
                ]
            ]
        ],
        3
    )
in
    Value.ReplaceType(Main.Func, Main.Type meta Main.Docs)