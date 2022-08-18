Attribute VB_Name = "modQueriesTables"
'#######################################################################################
' modQueriesTables
'
' VBA procedures related to refreshing Power Query M queries, and working with 
' ListObjects (tables)
'
'#######################################################################################

Option Explicit

Public Sub RefreshQueriesFromListObject( _
    ByRef tbRefresh As ListObject, _
    Optional ByVal UserPromptMsg As String = vbNullString, _
    Optional ByVal UpdateCompleteMsg As String = vbNullString, _
    Optional ByRef updatePivot As Variant = vbNullString _
)
'#######################################################################################
' RefreshQueriesFromListObject
'
' Refreshes queries listed in the given table, in the order they appear.
'
' Args:
'   tbRefresh: ListObject
'       a one-column wide table containing the names of queries in data rows
'   UserPromptMsg: string
'       Optional; overrides the default confirmation prompt
'   UpdateCompleteMsg: string
'       Optional; overrides the default completion message
'   updatePivot: boolean
'       Optional; sets whether all pivot tables will also be updated following the query
'       refreshes. When omitted, prompts the user for confirmation.
'
' Affects:
'   All listed queries will be refreshed, which will cause loaded tables to update
'   as well as all calculations that depend on them.
'#######################################################################################

Dim wb As Workbook
Set wb = ThisWorkbook
wb.Queries.FastCombine = True

Dim refreshList() As Variant
If tbRefresh.ListColumns(1).DataBodyRange.Count = 1 Then
    refreshList = Array(tbRefresh.ListColumns(1).DataBodyRange.Value)
Else
    refreshList = tbRefresh.ListColumns(1).DataBodyRange.Value
End If

Dim iterVar As Variant
Dim iterSheet As Worksheet

'Check the refreshList for any invalid query names. exit sub if any found
Dim validConnections As Collection
Set validConnections = New Collection

Dim invalidConnections As String
Dim validAsIs As Boolean
Dim validQueryPrefix As Boolean

invalidConnections = vbNullString
validAsIs = False
validQueryPrefix = False

Dim iterCon As Variant
Dim checkConnection As Variant
For Each iterCon In refreshList
    On Error Resume Next
    checkConnection = wb.Connections(iterCon)
    validAsIs = (Err.Number = 0)
    Err.Clear
    
    On Error Resume Next
    checkConnection = wb.Connections("Query - " & iterCon)
    validQueryPrefix = (Err.Number = 0)
    Err.Clear
    
    If validAsIs Then
        validConnections.Add iterCon
    ElseIf validQueryPrefix Then
        validConnections.Add "Query - " & iterCon
    Else
        invalidConnections = invalidConnections & vbNewLine & iterCon
    End If
Next iterCon

If Len(invalidConnections) > 0 Then
    MsgBox ( _
        "The following connections were not found. No queries have been refreshed." & _
        vbNewLine & _
        vbNewLine & _
        "Ensure these queries exist and their name is consistent with how its written in Queries & Connections." & _
        vbNewLine & _
        invalidConnections _
    )
    Exit Sub
End If

'prompt user to start all refreshes
Dim msgContinueResult As VbMsgBoxResult
Dim msgContinuePrompt As String
If IsMissing(UserPromptMsg) Or UserPromptMsg = vbNullString Then
    msgContinuePrompt = ( _
        "Update " & tbRefresh.DataBodyRange.Count & " queries? This will take several minutes. A message will notify you of completion." & _
        vbNewLine & _
        vbNewLine & _
        "Progress is tracked in the bottom left corner." _
    )
Else
    msgContinuePrompt = UserPromptMsg
End If
msgContinueResult = MsgBox(msgContinuePrompt, vbOKCancel)

' override completion message, if provided
Dim msgComplete As String
If IsMissing(UpdateCompleteMsg) Or UpdateCompleteMsg = vbNullString Then
    msgComplete = "Data Refreshes Complete"
Else
    msgComplete = UpdateCompleteMsg
End If

' prompt user to start refreshes
If msgContinueResult = vbOK Then
    Dim progressCurrent As Integer
    Dim progressMax As Integer
    Dim PauseTime, StartTime As Single
    
    PauseTime = 1 'adding 1 second pause between queries to allow for cancelling macro during runtime
    progressCurrent = 1
    progressMax = validConnections.Count
    
    Application.Calculation = xlCalculationManual
    For Each iterCon In validConnections
        Application.StatusBar = "Refreshing: " & iterCon & " | Progress: " & progressCurrent & "/" & progressMax
        wb.Connections(iterCon).OLEDBConnection.BackgroundQuery = False
        wb.Connections(iterCon).OLEDBConnection.Refresh
        progressCurrent = progressCurrent + 1
        StartTime = Timer
        Do While Timer < StartTime + PauseTime
            DoEvents
        Loop
    Next iterCon
    Application.Calculation = xlCalculationAutomatic
    
    Application.StatusBar = False

    'update pivot tables
    If updatePivot = vbNullString Then
        RefreshAllPivotTables True, UpdateCompleteMsg
    ElseIf updatePivot = True Then
        RefreshAllPivotTables False, UpdateCompleteMsg
    Else
        MsgBox msgComplete
    End If
        
End If

End Sub

Public Sub RefreshAllPivotTables( _
    Optional ByVal PromptRefresh As Boolean = True, _
    Optional ByVal msgComplete As String = vbNullString)
'#######################################################################################
' RefreshAllPivotTables
'
' Refresh all pivot tables in ThisWorkbook
'
' Args:
'   PromptRefresh: boolean (Default true)
'       Optional; sets whether user will be prompted to confirm the pivot table refreshes
'       Default True
'   msgComplete: string
'       overrides the default message sent when the refresh are complete
'
' Affects:
'   All pivot tables in ThisWorkbook will be updated to the latest data in their source.
'#######################################################################################

Dim wb As Workbook
Set wb = ThisWorkbook

' check if wb has any pivot tables
Dim countPivot As Long
Dim iterSheet As Worksheet
countPivot = 0
For Each iterSheet In wb.Sheets
    countPivot = countPivot + iterSheet.PivotTables.Count
Next iterSheet

' override default complete message, if provided
Dim msgCompleteDefault As String
If countPivot = 0 Then
    msgCompleteDefault = "Data Refreshes Complete"
ElseIf countPivot = 1 Then
    msgCompleteDefault = "Complete. " & countPivot & " pivot table in this workbook has been refreshed"
Else
    msgCompleteDefault = "Complete. " & countPivot & " pivot tables in this workbook have been refreshed"
End If

If IsMissing(msgComplete) Or msgComplete = vbNullString Then
    msgComplete = msgCompleteDefault
End If

' prompt user to update all pivot tables or not
Dim msgPivotPrompt As String
Dim msgPivotResult As VbMsgBoxResult
Dim updatePivot As Boolean
If countPivot = 0 Then
    ' do not prompt if there are no pivot tables
    updatePivot = False
ElseIf PromptRefresh Then
    ' check with user if there are pivot tables
    msgPivotPrompt = ( _
        "There are " & countPivot & " pivot tables in this workbook, which may or may not be connected to recently updated data connections." & _
        "Would you like to refresh all pivot tables?" _
    )
    msgPivotResult = MsgBox(msgPivotPrompt, vbYesNo)
    
    If msgPivotResult = vbYes Then
        updatePivot = True
    Else
        updatePivot = False
    End If
Else
    'if PromptRefresh is false, skip the user prompt and update all pivot tables
    updatePivot = True
End If

' refresh all pivot tables, output when complete
Dim iterPivot As PivotTable
If updatePivot Then
    For Each iterSheet In wb.Sheets
        For Each iterPivot In iterSheet.PivotTables
            iterPivot.RefreshTable
        Next iterPivot
    Next iterSheet
End If

MsgBox msgComplete

End Sub

Public Function CreateSortOperation(ByVal ColumnName As String, ByVal SortOrder As XlSortOrder) As Variant
'#######################################################################################
' CreateSortOperation -> Variant
'
' Creates a 2-item array for use in SortTable() SortOperations collection
'
' Args:
'   ColumnName: string
'       name of the column to use as sortkey
'   SortOrder: XLSortOrder
'       Ascending or Descending using the XLSortOrder enum
'
' Returns:
'   (1 to 2) indexed array of (column name, XLSortOrder)
'#######################################################################################

Dim oArray(1 To 2) As Variant
oArray(1) = ColumnName
oArray(2) = SortOrder

CreateSortOperation = oArray

End Function

Public Sub CopyData_TableToTable( _
    ByVal inTable As ListObject, _
    ByVal outTable As ListObject, _
    ByVal AllowColsToRight As Boolean, _
    Optional ByVal clear_inTable As Boolean = False, _
    Optional ByVal sort_operations_outTable As Collection)
'#######################################################################################
' CopyData_TableToTable
'
' Appends all rows of inTable to the bottom of outTable in new rows
'
' All column names in inTable must exist in outTable, in the same order, without
' additional columns between them. AllowColsToRight optional arg can be used to permit
' extra columns that do not interfere with the copy/paste from inTable.
'
' Args
'   inTable: ListObject
'       the table to copy from
'   outTable: ListObject
'       the table to append onto
'   AllowColsToRight: boolean
'       sets whether execution will stop if outTable has additional columns to the right
'       of required columns
'   clear_inTable: boolean
'       Optional; Default False; sets whether inTable will have all rows deleted after
'   sort_operations_outTable: collection
'       Optional; performs a sort on the output after pasting new rows
'#######################################################################################

Dim sameHeaders As Boolean
sameHeaders = False

'check that inTable is not empty
If inTable.DataBodyRange Is Nothing Then
    MsgBox ("modQueriesTables.CopyData_TableToTable() was called on an empty table. " & _
            "The process did not complete." & vbNewLine & vbNewLine & TableInfo(inTable) _
    )
    Exit Sub
End If

'check that both tables have the same number of columns with the same names, and in the same order
Dim i As Integer
If (inTable.HeaderRowRange.Count = outTable.HeaderRowRange.Count) Or AllowColsToRight Then
    sameHeaders = True
    For i = 1 To inTable.HeaderRowRange.Count Step 1
        If inTable.HeaderRowRange(1, i) <> outTable.HeaderRowRange(1, i) Then
            sameHeaders = False
        End If
    Next i
End If

'copy/paste inTable data values to the bottom of outTable
If sameHeaders Then
    inTable.DataBodyRange.Copy
    outTable.HeaderRowRange(1, 1).Offset(outTable.ListColumns(1).DataBodyRange.Count + 1, 0).PasteSpecial xlPasteValues
    Application.CutCopyMode = False
Else
    MsgBox ("modQueriesTables.CopyData_TableToTable() was called on tables without matching columns. The process did not complete." & _
        vbNewLine & vbNewLine & _
        "-- SOURCE TABLE --" & vbNewLine & _
        TableInfo(inTable) & _
        vbNewLine & vbNewLine & _
        "-- DEST TABLE --" & vbNewLine & _
        TableInfo(outTable) _
    )
End If
    
'delete inTable data per clear_inTable argument
If clear_inTable Then
    inTable.DataBodyRange.Delete
End If

'sort outTable if any sort operations were provided
If Not (IsMissing(sort_operations_outTable) Or sort_operations_outTable Is Nothing) Then
    SortTable outTable, sort_operations_outTable
End If

    
End Sub

Private Function TableInfo(ByVal tb As ListObject) As String
'#######################################################################################
' TableInfo -> string
'
' Returns a string with location info for the given ListObject
'
' Args:
'   tb: listObject
'       The table to describe
'
' Returns:
'   Multi-line string of format:
'       Workbook:   {WorkbookName}
'       Worksheet:  {WorksheetName}
'       Table:      {TableName}
'#######################################################################################

TableInfo = ("Workbook: " & vbTab & tb.Parent.Parent.Name & vbNewLine & _
    "Worksheet: " & vbTab & tb.Parent.Name & vbNewLine & _
    "Table: " & vbTab & vbTab & tb.Name _
    )

End Function

Public Sub SortTable(ByVal TableToSort As ListObject, ByVal SortOperations As Collection)
'#######################################################################################
' SortTable
'
' Sorts the given table with the operations provided in SortOperations collection
'
' Column names within SortOperations that do not exist are ignored.
' If SortOperations is an empty collection, or no columns listed exist in the table,
' then all existing sort fields are cleared.
'
' Args:
'   TableToSort: ListObject
'       Input table to be sorted in place
'   SortOperations: Collection
'       Collection of SortOperation arrays where:
'           A(1) = ColumnName as String and A(2) = SortOrder as xlSortOrder
'       Recommend creating this via CreateSortOperations() calls added to a collection
'          Coll.Add CreateSortOperation("Column Name", xlAscending)
'          Coll.Add CreateSortOperation("Second Level", xlDescending)
'
' Affects:
'   Sorts TableToSort in place
'
' Raises:
'   Error 457: This key is already associated with an element of this collection
'#######################################################################################

'late binding Scripting.Dictionary so that no reference is needed
'the dictionary is used to generate an error if a column is reused in SortOperations
Dim dict As Object
Set dict = CreateObject("Scripting.Dictionary") 
Dim i As Integer
For i = 1 To SortOperations.Count
    dict.Add SortOperations(i)(1), SortOperations(i)(2)
Next i

Dim AllColumnsExist As Boolean
AllColumnsExist = True

Dim MissingColumns As String
MissingColumns = vbNullString

Dim k As Variant
Dim rng As Range
For Each k In dict.keys
    With TableToSort.HeaderRowRange
        Set rng = .Find(What:=k, After:=.Cells(.Cells.Count), MatchCase:=False)
        If rng Is Nothing Then
            dict.Remove k
            MissingColumns = MissingColumns & vbNewLine & k
            AllColumnsExist = False
        End If
    End With
Next k

If Not AllColumnsExist Then
    MsgBox ("modQueriesTables.SortTable() tried to sort on column names that did not exist in the table. " & _
        vbNewLine & vbNewLine & _
        "The sort operation continued by ignoring them. " & _
        vbNewLine & vbNewLine & _
        "To prevent this message, rename the columns in the table or in the macro code to match each other." & _
        vbNewLine & vbNewLine & _
        TableInfo(TableToSort) & vbNewLine & _
        "Missing Columns: " & _
        MissingColumns & _
        vbNewLine & vbNewLine & _
        "Macro execution will continue after pressing OK.")
End If

If SortOperations.Count = 0 Then
    TableToSort.Sort.SortFields.Clear
Else
    With TableToSort.Sort
        .SortFields.Clear
        
        For Each k In dict
            .SortFields.Add2 Key:=TableToSort.ListColumns(k).DataBodyRange, _
                SortOn:=xlSortOnValues, Order:=dict(k), DataOption:=xlSortNormal
        Next k
        
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End If

End Sub