/* 
CurrentWorkbook.NamedRange

Return the value referenced by the workbook scoped named range. The return type is inferred from the structure of
[Content] for the named range.

Args
----
name: The named range to lookup, by default will only pull workbook scoped names.
sheet: [Optional] Name of the worksheet where the named range is scoped.
forceTable: Optional. Use true to always return the Content table accessed by Excel.CurrentWorkbook() function 
    instead of inferring the return type. Using forceTable=true is equivalent to writing 
    Excel.CurrentWorkbook(){[Name=name]}[Content]

ReturnType
----------
+ A literal if Content is a 1x1 table with "Column1" as the only column
+ A list if Content is a 1xN table with "Column1" as the only column
+ A table if neither of the above is true, or forceTable=true
*/
let
    AllNames = Excel.CurrentWorkbook()[Name],
    
    Main.Func = (name as text, optional sheet as nullable text, optional forceTable as nullable logical) as any =>
        let
            _Name = if sheet = null then name else Text.Format("#[sh]![nm]", [sh=sheet, nm=name]),
            TryContent = try Excel.CurrentWorkbook(){[Name=_Name]}[Content],
            Content =
                if TryContent[HasError] then
                    error
                        Error.Record(
                            "Argument Error",
                            Text.Format(
                                "'#[nm]' is not a valid named range for this workbook.",
                                [nm = name]
                            ),
                            Text.Combine(
                                {
                                    "Check that the spelling is correct, and that the named range is scoped to the",
                                    "workbook. Worksheet scoped named ranges can still be accessed using the optional",
                                    "sheet parameter."
                                },
                                " "
                            )
                        )
                else
                    TryContent[Value],
            ReturnLiteral =
                Table.RowCount(Content)
                = 1
                and Table.ColumnNames(Content)
                = {
                    "Column1"
                },
            ReturnList =
                Table.RowCount(Content)
                > 1
                and Table.ColumnNames(Content)
                = {
                    "Column1"
                },
            Output =
                if forceTable = true then
                    Content
                else if ReturnLiteral then
                    Content{0}[Column1]
                else if ReturnList then
                    Content[Column1]
                else
                    Content
        in
            Output,
    Main.Type = Type.ForFunction(
        [
            ReturnType = type any,
            Parameters = [
                name = type text meta [
                    Documentation.FieldCaption = "name: named range",
                    Documentation.AllowedValues = AllNames
                ],
                optional sheet = type nullable text meta [
                    Documentation.FieldCaption = "sheet: worksheet scope"
                ],
                optional forceTable = type nullable logical meta [
                    Documentation.FieldCaption = "forceTable: always return content as a table"
                ]
            ]
        ],
        Type.FunctionRequiredParameters(Value.Type(Main.Func))
    ),
    Main.Docs = [
        Documentation.Name = "CurrentWorkbook.NamedRange",
        Documentation.LongDescription = Text.Combine(
            {
                "Return the value referenced by the workbook named range.",
                "Worksheet scoped ranges can be accessed by also providing the 'sheet' parameter.<br><br>",
                
                "<b>Args</b><br>-----<br>",
                "<b>name</b> as text: named range<br>",
                "<b>sheet</b> as nullable text: [Optional] sheet name for worksheet scoped named ranges<br>",
                "<b>forceTable</b> as nullable logical: [Optional] always return a table of the named range",
                "instead of inferring the return type<br><br>",
                
                "<b>Returns</b><br>---------<br>Values are assumed to be:<br>",
                "+ A literal if Content is a 1x1 table with 'Column1'<br>",
                "+ A list if Content is a 1xN table with 'Column1'<br>",
                "+ A table if neither of the above is true<br><br>",
                "The table return format is the default for Excel.CurrentWorkbook()[Content] records.",
                "forceTable=true may optionally be used to return the named range as a table regardless of the layout",
                "structure - this is equivalent to writing Excel.CurrentWorkbook(){[Name=name]}[Content].<br><br>",
                
                "<b>Raises</b><br>-------<br>",
                "Argument Error: the named range does not exist in the workbook scope"
            },
            " "
        ),
        Documentation.Examples = {
            [
                Description = "Name refers to a single cell",
                Code = "CurrentWorkbook_NamedRange(""TheNumberFour"")",
                Result = "4"
            ],
            [
                Description = "Name refers to a single cell, forceTable=true",
                Code = "CurrentWorkbook_NamedRange(""TheNumberFour"", null, true)",
                Result = "Column1#(lf)-------#(lf)4"
            ],
            [
                Description = "Name refers to a range of cells",
                Code = "CurrentWorkbook_NamedRange(""Alphabet"")",
                Result = "{""A"",""B"",""C"",""D"", ..., ""Z""}"
            ],
            [
                Description = "Name refers to a range of cells, forceTable=true",
                Code = "CurrentWorkbook_NamedRange(""Alphabet"", null, true)",
                Result = "Column1#(lf)-------#(lf)A#(lf)B#(lf)C#(lf)..#(lf)Z"
            ],
            [
                Description = "Name refers to a table only scoped to sheet 'abc'",
                Code = "CurrentWorkbook_NamedRange(""AlphabetWithPosition"", ""abc"")",
                Result = Text.Combine(
                    {
                        "Pos | Letter#(lf)",
                        "------------#(lf)",
                        "  1 |  A#(lf)",
                        "  2 |  B#(lf)",
                        "  3 |  C#(lf)",
                        " .. | ..#(lf)",
                        " 26 |  Z"
                    }
                )
            ]
        }
    ]
in
    Value.ReplaceType(Main.Func, Main.Type meta Main.Docs)