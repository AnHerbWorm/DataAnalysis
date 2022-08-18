Attribute VB_Name = "modErrorHandling"
'#######################################################################################
' modErrorHandling
'
' Error handling module to reraise errors through the call stack
'
' Credit:
'   https://excelmacromastery.com/vba-error-handling/#A_Complete_Error_Handling_Strategy
'   https://www.youtube.com/watch?v=lR5e8gyA69U
'#######################################################################################

Option Explicit

Dim AlreadyUsed As Boolean

' Reraises an error and adds line number and current procedure name
Sub RaiseError(ByVal errorNo As Long _
                , ByVal src As String _
                , ByVal proc As String _
                , ByVal desc As String)

    Dim sSource As String

    ' If called for the first time then add line number
    If AlreadyUsed = False Then
        
        ' Add procedure to source
        sSource = sSource & vbNewLine & proc
        AlreadyUsed = True
        
    Else
        ' If error has already been raised simply add on procedure name
        sSource = src & vbNewLine & proc
    End If
    
    ' Pause the code here when debugging
    '(To Debug: "Tools->VBA Properties" from the menu.
    ' Add "Debugging=1" to the     ' "Conditional Compilation Arguments.)
#If Debugging = 1 Then
    Debug.Assert False
#End If

    ' Reraise the error so it will be caught in the caller procedure
    ' (Note: If the code stops here, make sure DisplayError has been
    ' placed in the topmost procedure)
    Err.Raise errorNo, sSource, desc

End Sub

' Displays the error when it reaches the topmost sub
' Note: You can add a call to logging from this sub
Sub DisplayError(ByVal src As String, ByVal desc As String _
                    , ByVal sProcname As String)

    ' Check If the error happens in topmost sub
    If AlreadyUsed = False Then
        ' Reset string to remove "VBAProject"
        src = vbNullString
    End If

    ' Build the final message
    Dim sMsg As String
    sMsg = "The following error occurred: " & vbNewLine & Err.Description _
                    & vbNewLine & vbNewLine & "Error Location is: "
    sMsg = sMsg & src & vbNewLine & sProcname

    ' Display the message
    MsgBox sMsg, Title:="Error"

    ' reset the boolean value
    AlreadyUsed = False

End Sub
