Attribute VB_Name = "modFileHandling"
'#######################################################################################
' modFileHandling
'
' VBA procedures making use of Scripting.FileSystemObject and file selection dialogs
'
'#######################################################################################

Option Explicit

Public Function SelectExcelFilesToCollection( _
    Optional ByVal UserPrompt As String = "", _
    Optional ByVal StartFolder As String = "", _
    Optional ByVal SelectMultiple As Boolean = False, _
    Optional ByVal NoFileMsg As String = "", _
    Optional ByVal SuppressNoFileMsg As Boolean = False _
    ) As Collection
'#######################################################################################
' SelectExcelFilesToCollection
'
'   Open a file selection dialog to select one or more excel files and return their
'   file paths in a collection object
'
' Args:
'   UserPrompt: string
'       Override default message for the initial popup confirmation window
'   StartFolder: string (Default ThisWorkbook.Path if omitted)
'       folder path to open the file selection to initially
'   SelectMultiple: bool (Default false)
'       whether more than 1 filepath may be returned
'   NoFileMsg: string
'       Override default message for when no file is returned
'   SuppressNoFileMsg: bool (Default false)
'       Turn off the MsgBox interruption when no file is selected. Intended for use when
'       the procedure calling this function has it's own handling for this case
'
' Returns:
'   A collection of strings, where each item is a file path.
'   If no file is selected an empty collection is returned.
'
'NOTE:
'   Any procedure calling this one should check that the output collection
'   has one or more elements before proceeding
'#######################################################################################

Dim SelectedFiles As Collection
Set SelectedFiles = New Collection

'-- Check that the supplied folder exists, use default if not
Dim fso As Object
Set fso = CreateObject("Scripting.FileSystemObject")

Dim DefaultPath As String
If fso.FolderExists(StartFolder) Then
    DefaultPath = ThisWorkbook.Path
Else
    DefaultPath = StartFolder
End If

'-- Set default messages
Dim mboxPrompt As String
Dim mboxSelection As VbMsgBoxResult

If UserPrompt = "" Then
    If SelectMultiple Then
        mboxPrompt = "Select one or more Excel files"
    Else
        mboxPrompt = "Select an Excel file"
    End If
Else
    mboxPrompt = UserPrompt
End If

Dim mboxNoFileMsg As String
If NoFileMsg = "" Then
    mboxNoFileMsg = "No file was selected. The subroutine that called this function may fail " & _
        "if it is not setup to handle 0 length collections." & _
        vbNewLine & vbNewLine & _
        "If the sub is setup to handle 0 length collections it is recommended to change or suppress " & _
        "this message box"
Else
    mboxNoFileMsg = NoFileMsg
End If

'-- Prompt user to start the macro
mboxSelection = MsgBox(mboxPrompt, vbOKCancel)
If mboxSelection = vbCancel Then
    GoTo proc_exit_nofile
End If

'-- Open file selection window and get file paths
Dim iter As Integer
With Application.FileDialog(msoFileDialogOpen)
    .InitialFileName = StartFolder
    .AllowMultiSelect = SelectMultiple
    .Filters.Add "Excel Files", "*.xlsx; *.xlsm; *.xls; *.xlsb"
    .Show
    If .SelectedItems.Count = 0 Then
        GoTo proc_exit_nofile
    Else
        For iter = 1 To .SelectedItems.Count
            SelectedFiles.Add .SelectedItems(iter)
        Next iter
    End If
End With

'-- Exit routines
proc_exit:
    Set SelectExcelFilesToCollection = SelectedFiles
    Exit Function

proc_exit_nofile:
    If Not SuppressNoFileMsg Then
        MsgBox mboxNoFileMsg, vbInformation
    End If
    GoTo proc_exit

End Function

Public Sub CreateFolderWithParentDirs(ByVal folderPath As String, Optional ByVal hideCompleteMsg As Boolean = False, optional ByVal hideErrorMsg as Boolean = False)
'#######################################################################################
' CreateFolderWithParentDirs
'
'   Creates the folder and all parent directories that do not exist. Assumes that the
'   user has sufficient privileges to create folders in the root directory.
'
' Args:
'   folderPath: string
'       The directory to create.
'   hideCompleteMsg: optional bool (Default: False)
'       Sets whether to hide the message box confirming folder creation
'   hideErrorMsg: optional bool (Default: False)
'       Sets whether to hide the message box displaying failure information
'
' Affects:
'   The file system as this sub will create directories as needed.
'#######################################################################################
Dim fso As Object
Set fso = CreateObject("Scripting.FileSystemObject")

'-- Find missing folders
Dim currentFolder As String
Dim existingRoot As String

Dim toCreate As Collection
Set toCreate = New Collection

currentFolder = folderPath
Do While currentFolder <> vbNullString 
    ' GetParentFolderName returns vbNullString when it goes higher than the drive level
    ' so we can break the loop on this condition
    If fso.FolderExists(currentFolder) Then
        existingRoot = currentFolder
        Exit Do
    Else
        toCreate.Add currentFolder
    End If
    currentFolder = fso.GetParentFolderName(currentFolder)
Loop

'-- Creates all folders, starting with the lowest level
Dim driveExists as Boolean
driveExists = (currentFolder <> vbNullString)

If driveExists Then
    For depth = toCreate.count To 1 Step -1
        fso.CreateFolder (toCreate(depth))
    Next depth
Else
    If Not hideErrorMsg then
        MsgBox "Script Failure <modFileHandling.CreateFolderWithParentDirs>: " _ 
            & vbNewLine & vbNewLine
            & "Cannot create a drive. Verify the folderPath is correct and/or that drive '" & fso.GetDriveName(folderPath) & "' is accessible." _ 
            & vbNewLine & vbNewLine _ 
            & "folderPath = " & folderPath 
    End if
End If

'-- Optionally confirm completion
If Not hideCompleteMsg Then
    If existingRoot = folderPath Then
        MsgBox "Script Complete <modFileHandling.CreateFolderWithParentDirs>: " _ 
            & vbNewLine & vbNewLine _
            & "'" & folderPath & "' already existed. No action was taken."
    Else
        MsgBox "Script Complete <modFileHandling.CreateFolderWithParentDirs>: " _
            & vbNewLine & vbNewLine _
            & "Created '" & folderPath & "' by making" & Str(toCreate.count) & " folders starting in '" & existingRoot & "'."
    End If
End If

End Sub
