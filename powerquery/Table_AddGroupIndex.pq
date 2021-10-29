/* 
Table.AddGroupIndex

Chains Table.ExpandTableColumn(Table.AddIndexColumn(Table.Group)) functions to append a new index column
identifying which group each row in the source table belongs to.
*/
let
    Main.Func = (table as table, key as anynonnull, newColumnName as text) as table =>
        Table.ExpandTableColumn(
            Table.AddIndexColumn(
                Table.Group(table, key, {{"rows", each _, Value.Type(table)}}),
                newColumnName
            ),
            "rows",
            List.RemoveItems(Table.ColumnNames(table), try key otherwise {key})
    ),
    Main.Type = Type.ForFunction(
        [
            ReturnType = type table,
            Parameters = [
                table = type table meta [
                    Documentation.FieldCaption = "source table"
                ],
                key = type anynonnull meta [
                    Documentation.FieldCaption = "column(s) to group on"
                ],
                newColumnName = type text meta [
                    Documentation.FieldCaption = "column name for index"
                ]
            ]
        ],
        Type.FunctionRequiredParameters(Value.Type(Main.Func))
    ),
    Main.Docs = [
        Documentation.Name = "Table.AddGroupIndex",
        Documentation.LongDescription = Text.Combine(
            {
                "Appends <code>newColumnName</code> to <code>table</code> where each index value identifies",
                "which group that record belongs to after grouping on <code>key</code>.",
                "Indexes start at 0, increment by 1, and are numbered in the order observed.<br><br>",
                
                "<b>Args:</b><br>",
                "<code>table</code>: table to group and index<br>",
                "<code>key</code>: column name(s) to group on<br>",
                "<code>newColumnName</code>: name of the index column appended<br><br>",
                
                "<b>Returns:</b><br>",
                "The source table with one index column appended. The index column starts at 0, increments by 1",
                ", and groups are numbered in the order they are observed."
            },
            " "
        ),
        Documentation.Examples = {
            [
                Description = "Basic example",
                Code = Text.Combine(
                    {
                        "Table_AddGroupIndex(#(lf)",
                        "  #table(type table#(lf)",
                        "    [One=text, Two=text, Val=number],#(lf)",
                        "    {#(lf)",
                        "      {""A"", ""A"", 1},#(lf)",
                        "      {""A"", ""A"", 2},#(lf)",
                        "      {""A"", ""B"", 1},#(lf)",
                        "      {""B"", ""B"", 5}#(lf)",
                        "    }#(lf)",
                        "  ),#(lf)",
                        "  {""One"", ""Two""},#(lf)",
                        "  ""OneTwoKey""#(lf))"
                    }
                ),
                Result = Text.Combine(
                    {
                        "#table(type table#(lf)",
                        "  [One=text, Two=text, Val=number, OneTwoKey=number],#(lf)",
                        "  {#(lf)",
                        "    {""A"", ""A"", 1, 0},#(lf)",
                        "    {""A"", ""A"", 2, 0},#(lf)",
                        "    {""A"", ""B"", 1, 1},#(lf)",
                        "    {""B"", ""B"", 5, 2}#(lf)",
                        "  }#(lf)",
                        ")"
                    }
                )
            ]
        }
    ]
in
    Value.ReplaceType(Main.Func, Main.Type meta Main.Docs)