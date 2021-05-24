/*
List.Combinations

Generate a list of non-repeating combinations of length combinationSize from the set
This will create a number of subsets equal to the result of:
--Number.Combinations(List.Count(set), combinationSize)
--nCr from mathematics notation
*/
let
    /* Validation Checks */
    Param.CombSize = (combinationSize, set) =>
        if combinationSize >= 1 and combinationSize <= List.Count(set) then
            combinationSize
        else
            error Error.Record(
                "Parameter Error",
                "combinationSize must be between 1 and the length of set",
                [Value=combinationSize, Min=1, Max=List.Count(set)]
            ),
    Param.Set = (set) => if List.IsDistinct(set) then set else Error.Record(
        "Parameter Error",
        "set must be distinct to provide non-repeating combinations",
        [IsDistinct=false]
    ),
    /* Main Function */
    Main.Func =
        (set as list, combinationSize as number) as list =>
            let
                CombSize = Param.CombSize(combinationSize, set),
                Set = Param.Set(set),

                MakeSubsets = (n, r, optional tb) =>
                    /*
                    Recursive function adds one column to the table with new indices based on the column count. New columns are lists which get expanded into new records for the next iteration.

                    For example, if doing nCr = 4C2 = 6 combinations, the recursive calls will stop after the 2nd column is added
                    
                    The first iteration will create:
                        #table({"col0"},{{0}, {1}, {2}})
                    
                    The second iteration adds a column of filtered lists:
                        #table({"col0", "col1"},{
                            {0,{1,2,3}}, {1,{2,3}, {2,{3}}
                        }) which is expanded to the cartesian-product of the first column and it's nested list in the second column
                        
                    To finally output a table of 6 rows x 2 columns:
                    #table({"col0", "col1"}, {
                            {0, 1}, {0, 2}, {0, 3},
                            {1, 2}, {1, 3},
                            {2, 3}
                        }) 
                    */
                    let
                        // iteration and index partition setup. The iteration is determined by the number of columns in the table, and recursion stops when the column count is equal to r
                        CurrentIteration = try Table.ColumnCount(tb) otherwise 0,
                        Partition = {
                            CurrentIteration
                            ..
                            CurrentIteration + n - r
                        },
                        // column names
                        PrevCol = "c" & Number.ToText(CurrentIteration - 1),
                        CurrCol = "c" & Number.ToText(CurrentIteration),
                        // initialization output; create a single column table of all the indices in IndexLists{0}
                        init =
                            Table.FromList(
                                Partition,
                                Splitter.SplitByNothing(),
                                {CurrCol},
                                null,
                                ExtraValues.Error
                            ),
                        // continuation output; add a new column with filtered lists then expand it
                        cont =
                            Table.ExpandListColumn(
                                Table.AddColumn(
                                    tb,
                                    CurrCol,
                                    (record) =>
                                        List.Select(
                                            Partition,
                                            each _ > Record.Field(record, PrevCol)
                                        )
                                ),
                                CurrCol
                            ),
                        // final output is the last continuation iteration (r - 1)
                        out =
                            if CurrentIteration = r then
                                tb
                            else if CurrentIteration = 0 then
                                @MakeSubsets(n, r, init)
                            else
                                @MakeSubsets(n, r, cont)
                    in
                        out,
                SubsetsAsList =
                    List.Transform(
                        Table.ToRows(MakeSubsets(List.Count(Set), CombSize)),
                        // each sublist has its indices converted back to values from the set
                        (sublist) => List.Transform(sublist, each Set{_})
                    )
            in
                SubsetsAsList,
    Main.Type = Type.ForFunction(
        [
            ReturnType = type list,
            Parameters = [
                set = type list meta [
                    Documentation.FieldCaption = "the set of values to get combinations of"
                ],
                combinationSize = type number meta [
                    Documentation.FieldCaption = "the size of each combination subset"
                ]
            ]
        ],
        2
    ),
    Main.Docs = [
        Documentation.Name = "List.Combinations",
        Documentation.LongDescription = "Generate a list of non-repeating combination subsets of length 'combinationSize' from the set.<br>The function will generate nCr # of combinations, which can be calculated in Power Query as Number.Combinations(List.Count(set), combinationSize)",
        Documentation.Examples = {
            [
                Description = "Subsets of 4 from ABCDE",
                Code = "List_Combinations({""A"", ""B"", ""C"", ""D"", ""E""}, 4)",
                Result = "{#(lf)  {""ABCD""},#(lf)  {""BCDE""}#(lf)}"
            ],
            [
                Description = "Subsets of 2 from 1234",
                Code = "List_Combinations({1, 2, 3, 4}, 2)",
                Result = "{#(lf)  {1,2},#(lf)  {1,3},#(lf)  {1,4},#(lf)  {2,3},#(lf)  {2,4},#(lf)  {3,4}#(lf)}"
            ]
        }
    ]
in
    Value.ReplaceType(Main.Func, Main.Type meta Main.Docs)