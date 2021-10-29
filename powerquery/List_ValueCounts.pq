/*
List_ValueCounts

Return nested list of distinct values and their count of occurrences within the source list.
*/
let
    Main.Func = (list as list, optional outputAs as nullable text) as any =>
        let
            ListAsTable = #table(type table [Value=any], List.Transform(list, each {_})),
            Grouped = Table.Group(ListAsTable, {"Value"}, {{"Count", Table.RowCount, Int64.Type}}),
            Results = [
                table = Grouped,
                lists = Table.ToList(table, each _),
                records = Table.ToRecords(table)
            ],
            Format = if outputAs = null then "table" else Text.Lower(outputAs),
            Result = Record.FieldOrDefault(Results, Format, Grouped)
        in
            Result,
    Main.Type = Type.ForFunction(
        [
            ReturnType = type any,
            Parameters = [
                list = type list meta [
                    Documentation.FieldCaption = "List"
                ],
                outputAs = type nullable text meta [
                    Documentation.FieldCaption = "Return Counts As",
                    Documentation.AllowedValues = {"table", "lists", "records"}
                ]
            ]
        ],
        1
    ),
    Main.Docs = [
        Documentation.Name = "List_ValueCounts",
        Documentation.LongDescription = Text.Combine(
            {
                "Counts each occurrence of distinct values in the input list.<br><br>",
                "<b>Args<br>------</b><br>",
                "<i>list</i>: Source list for value counts.<br>",
                "<i>outputAs</i>: Text representation of how to return the value counts. Options are list/records/table. Default 'table' will be applied if omitted or a non-allowed value is entered. See examples for more details.<br>",
                "<b>Returns<br>---------</b><br>",
                "Nested list where the inner list has structure: item{0} = value and item{1} = count."
            },
            ""
        ),
        Documentation.Examples = {
            [
                Description = "Return as table (DEFAULT)",
                Code = "List_ValueCounts({0, 0, 1, 4, 4, 4, 4, 4, 7, 8})",
                Result = Text.Combine({
                    "#table(type table [Value=any, Count=number], ",
                    "{ {0, 2}, {1, 1}, {4, 5}, {7, 1}, {8, 1} })"},
                    "#(lf)"
                )
            ],
            [
                Description = "Return as list of lists",
                Code = "List_ValueCounts({0, 0, 1, 4, 4, 4, 4, 4, 7, 8}, ""lists"")",
                Result = "{ {0, 2}, {1, 1}, {4, 5}, {7, 1}, {8, 1} }"
            ],
            [
                Description = "Return as list of records",
                Code = "List_ValueCounts({0, 0, 1, 4, 4, 4, 4, 4, 7, 8}, ""records"")",
                Result = Text.Combine({
                    "{ [Value=0, Count=2], [Value=1, Count=1], [Value=4, Count=5],",
                    "  [Value=7, Count=1], [Value=8, Count=1] }"},
                    "#(lf)"
                )
            ]
        }
    ]
in
    Value.ReplaceType(Main.Func, Main.Type meta Main.Docs)