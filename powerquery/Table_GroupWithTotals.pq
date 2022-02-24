/**
* compute and combine Table.Group() calls on subsets of sourceTable with sub/grand totals added
*
* performs initial checks on parameter values to ensure the function will complete as it can be
* expensive to find out late in the calculation sequence that a combination fails
*
* @name Table.GroupWithTotals
* @categories table
* @license MIT (c) 2021 W Bohm
* @author https://github.com/AnHerbWorm
* @source https://github.com/AnHerbWorm/DataAnalysis/tree/main/powerquery
* @result table
*/
let
    /* Subset Functions
    A subset is a non-nullable list of subtotals, where each subtotal has length > 0

    Subset.Type -> type: type definition for reference only. It is not ascribed anywhere but describes the 
        list {record} structure
    Subset.Columns -> list: all column field values in the subset
    Subset.IsValid -> logical: column field values are distinct and all required columns are present
    Subset.ToTable -> table: sourceTable with subtotal application for each record in the subset.
        Can return a table with 0 rows if the subtotals do not have any overlap in the results of their selector 
        functions
    Subset.Combinations -> list: list of all combinations of subsets of length n
    */
    Subset.Type = type {[column = text, alias = text, selector = function]},
    Subset.Columns = (subset as list) as list => List.Transform(subset, each _[column]),
    Subset.IsValid = (subset as list, requiredColumns as nullable list) as logical =>
        if requiredColumns = null then
            List.IsDistinct(Subset.Columns(subset))
        else
            List.IsDistinct(Subset.Columns(subset))
            and List.MatchesAll(requiredColumns, each List.Contains(Subset.Columns(subset), _)),
    Subset.ToTable = (subset as list, iter as number, tbl as table) as table =>
        let
            Curr = subset{iter},
            Recur = iter < List.Count(subset),
            ContTable = Table.TransformColumns(
                Table.SelectRows(tbl, each Curr[selector](Table.Column(_, Curr[column]))),
                {
                    {Curr[column], each Curr[alias], type text}
                }
            )
        in
            if Recur then
                @Subset.ToTable(subset, iter + 1, ContTable)
            else
                tbl,
    Subset.Combinations = (s as list, n as number) as list =>
        if n = 0 then
            {{}}
        else if List.Count(s) = 0 then
            {}
        else
            List.Transform(@Subset.Combinations(List.Skip(s), n - 1), each {List.First(s)} & _)
            &
            @Subset.Combinations(List.Skip(s), n),
    // Validation Checks
    Param.KeyColumns = (keyColumns, subtotals) =>
        let
            NotKeys = List.Select(
                List.Transform(subtotals, each _{0}),
                each not List.Contains(keyColumns, _)
            ),
            Details = 
                if keyColumns = null then
                    [IsNull=true]
                else
                    [Not Keys=Text.Combine(List.Distinct(NotKeys), ", ")],
            Err = Error.Record(
                "Parameter Error",
                "keyColumns cannot be null and must contain all columns listed in subtotals",
                Details
            ),
            Param = 
                if List.Count(NotKeys) = 0 then
                    keyColumns
                else
                    error Err
        in
            Param,
    Param.Subtotals = (subtotals, keyColumns) =>
        let
            Columns = List.Transform(subtotals, each _{0}),
            Aliases = List.Transform(subtotals, each _{1}),
            AliasByColumn = List.Accumulate(
                List.Zip({Columns, Aliases}),
                // generating a record of each column used with an empty list value
                // so that these lists can be populated with each alias used for duplication checking
                Record.FromList(
                    List.Repeat({{}}, List.Count(List.Distinct(Columns))),
                    List.Distinct(Columns)
                ),
                (state, current) => Record.TransformFields(state, {current{0}, each _ & {current{1}}})
            ),
            ColumnsWithDupes = Table.SelectRows(Record.ToTable(AliasByColumn), (tb) => not List.IsDistinct(tb[Value])),
            Err = Error.Record(
                "Parameter Error",
                "each distinct subtotals column cannot reuse an alias",
                [column=ColumnsWithDupes[Name]{0}, alias="..."]
            ),
            Param = 
                if Table.RowCount(ColumnsWithDupes) = 0 then
                    subtotals
                else
                    error Err
            in
                Param,
    Param.TotalsOnly = (totalsOnlyColumns, subtotals) =>
        let
            NotSubtotal = List.Select(
                totalsOnlyColumns,
                each not List.Contains(List.Transform(subtotals, each _{0}), _)),
            Err = Error.Record(
                "Parameter Error",
                "totalsOnlyColumns must be null or only contain columns listed in subtotals",
                [Not Subtotals=Text.Combine(List.Distinct(NotSubtotal), ", ")]
            ),
            Param = 
                if totalsOnlyColumns = null or List.Count(NotSubtotal) = 0 then
                    totalsOnlyColumns
                else
                    error Err
        in
            Param,
    // Main Function
    Main.Func = (sourceTable as table,
                 keyColumns as list,
                 subtotals as list,
                 aggregatedColumns as list,
                 optional totalsOnlyColumns as nullable list
                ) as table =>
        let
            KeyColumns = Param.KeyColumns(keyColumns, subtotals),
            Subtotals = Param.Subtotals(subtotals, keyColumns),
            TotalsOnly = Param.TotalsOnly(totalsOnlyColumns, subtotals),
        
            SubtotalsAsRecords = List.Transform(Subtotals, each Record.FromList(_, {"column", "alias", "selector"})),
            // the max number of elements in a subset is equal to the number of columns where a total is applied
            // the min number of elements in a subset is equal to the number of totalsOnlyColumns columns, 
            // or 1 if totalsOnlyColumns is null
            MaxSubsetSize = List.Count(List.Distinct(Subset.Columns(SubtotalsAsRecords))),
            MinSubsetSize = List.Min({MaxSubsetSize, try List.Count(TotalsOnly) otherwise 1}),
            AllSubsets = List.Accumulate(
                {MinSubsetSize..MaxSubsetSize},
                {},
                (state, current) => state & Subset.Combinations(SubtotalsAsRecords, current)
            ),
            ValidSubsets = List.Buffer(List.Select(AllSubsets, each Subset.IsValid(_, TotalsOnly))),
            BufferedSource = Table.Buffer(sourceTable),
            CreateAndCombineSubsetTables = Table.Combine(
                List.Transform(ValidSubsets, each Subset.ToTable(_, 0, BufferedSource))
            ),
            // apply standard Table.Group()
            // in cases where TotalsOnly is null, we also need to compute measures on the base table
            BaseMeasures = Table.Group(
                BufferedSource,
                KeyColumns,
                aggregatedColumns
            ),
            CombinationMeasures = Table.Group(
                CreateAndCombineSubsetTables,
                KeyColumns,
                aggregatedColumns
            ),
            OutputMeasures =
                if TotalsOnly = null then
                    CombinationMeasures & BaseMeasures
                else
                    CombinationMeasures
        in
            OutputMeasures,
    Main.Type = Type.ForFunction(
        [
            ReturnType = type table,
            Parameters = [
                sourceTable = type table meta [
                    Documentation.FieldCaption = "table to group"
                ],
                keyColumns = type list meta [
                    Documentation.FieldCaption = "list of column names to group on"
                ],
                subtotals = type list meta [
                    Documentation.FieldCaption = "subtotals"
                ],
                aggregatedColumns = type list meta [
                    Documentation.FieldCaption = "aggregatedColumns"
                ],
                optional totalsOnlyColumns = type nullable list meta [
                    Documentation.FieldCaption = "totalsOnlyColumns"
                ]
            ]
        ],
        4
    ),
    Main.Docs = [
        Documentation.Name = "Table.GroupWithTotals",
        Documentation.LongDescription = Text.Combine(
            {
                "Create and combine the results of Table.Group() on all combinations of user-defined subsets",
                "of the table.<br><br>",
            
                "<b>Args</b><br>----<br>",
                "<code>sourceTable</code>: table to group<br>",
                "<code>keyColumns</code>: list of column names to group on<br>",
                "<code>subtotals</code>: nested list of 3-element lists where each inner list is the specifications",
                "for creating a total.<br>",
                "sublist{0} <i>column as text</i>: name of the column to base the subtotal on. Must be in keyColumns.<br>",
                "sublist{1} <i>alias as text</i>: value that refers to the subtotal. All values in the column get",
                "replaced by this alias after selection. An error is raised if a column attempts to use the same",
                "alias twice.<br>",
                "sublist{2} <i>selector as function</i>: function that returns true/false when applied to each column value.<br>",
                "<code>aggregatedColumns</code>: see M Documentation https://docs.microsoft.com/en-us/powerquery-m/table-group.",
                "Columns used in aggregate calculations should not also be used for subtotals. The query will likely",
                "complete but produce error values within the final table.<br>",
                "<code>totalsOnlyColumns</code>: [Optional] list of column names where only the subtotals will be returned,",
                "instead of the typical behaviour of Table.Group that groups on each distinct value within keyColumns.<br>",
                
                "<br><b>Dependencies</b><br>----------------<br>",
                "List.Combinations (as List_Combinations)"
            },
            " "
        ),
        Documentation.Examples = {
            [
                Description = "5 Largest Canadian Cities by Province with subtotals",
                Code = Text.Combine({
                    "Table_GroupWithTotals(#(lf)",
                    "  Table.FromRecords({#(lf)",
                    "    [City=""Toronto"" , Province=""ON"", Pop=2.73],#(lf)",
                    "    [City=""Montreal"", Province=""QC"", Pop=1.70],#(lf)",
                    "    [City=""Calgary"" , Province=""AB"", Pop=1.24],#(lf)",
                    "    [City=""Ottawa""  , Province=""ON"", Pop=0.93],#(lf)",
                    "    [City=""Edmonton"", Province=""AB"", Pop=0.93]}#(lf)",
                    "  ),#(lf)",
                    "  {""Province""},#(lf)",
                    "  {#(lf)",
                    "    {""Province"", ""Canada"", each true},#(lf)",
                    "    {""Province"", ""West"", (prov) => List.Contains({""BC"", ""AB"", ""SK"", ""MB""}, prov)}#(lf)",
                    "  },#(lf)",
                    "  {#(lf)",
                    "    {""cities"", Table.RowCount, Int64.Type},#(lf)",
                    "    {""avg pop"", each List.Average([Pop]), type number}#(lf)",
                    "  },#(lf)",
                    "  null#(lf)",
                    ")"
                }),
                Result = Text.Combine({
                    "#table(#(lf)",
                    "  {""Province"", ""cities"", ""avg pop""},#(lf)",
                    "  {#(lf)",
                    "    {""Canada"", 5, ""1.506""},#(lf)",
                    "    {""West"", 2, 1.085},#(lf)",
                    "    {""ON"", ""2"", 1.83}#(lf)",
                    "    {""QC"", 1, 1.70},#(lf)",
                    "    {""AB"", 2, 1.085}#(lf)",
                    "  }#(lf)",
                    ")"
                })
            ],
            [
                Description = "5 Largest Canadian Cities by Province, subtotals only",
                Code = Text.Combine({
                    "Table_GroupWithTotals(#(lf)",
                    "  Table.FromRecords({#(lf)",
                    "    [City=""Toronto"" , Province=""ON"", Pop=2.73],#(lf)",
                    "    [City=""Montreal"", Province=""QC"", Pop=1.70],#(lf)",
                    "    [City=""Calgary"" , Province=""AB"", Pop=1.24],#(lf)",
                    "    [City=""Ottawa""  , Province=""ON"", Pop=0.93],#(lf)",
                    "    [City=""Edmonton"", Province=""AB"", Pop=0.93]}#(lf)",
                    "  ),#(lf)",
                    "  {""Province""},#(lf)",
                    "  {#(lf)",
                    "    {""Province"", ""Canada"", each true},#(lf)",
                    "    {""Province"", ""West"", (prov) => List.Contains({""BC"", ""AB"", ""SK"", ""MB""}, prov)}#(lf)",
                    "  },#(lf)",
                    "  {#(lf)",
                    "    {""cities"", Table.RowCount, Int64.Type},#(lf)",
                    "    {""avg pop"", each List.Average([Pop]), type number}#(lf)",
                    "  },#(lf)",
                    "  {""Province""}#(lf)",
                    ")"
                }),
                Result = Text.Combine({
                    "#table(#(lf)",
                    "  {""Province"", ""cities"", ""avg pop""},#(lf)",
                    "  {#(lf)",
                    "    {""Canada"", 5, ""1.506""},#(lf)",
                    "    {""West"", 2, 1.085}#(lf)",
                    "  }#(lf)",
                    ")"
                })
            ]
        }
    ]
in
    Value.ReplaceType(Main.Func, Main.Type meta Main.Docs)