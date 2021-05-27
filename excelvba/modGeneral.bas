Attribute VB_Name = "modGeneral"
'#######################################################################################
' modGeneral
'
' General purpose VBA procedures that may be useful for any project
'
' Module Variables:
'   UpdateState: enum
'       provide intellisense for App_UpdateStatus() calls
'   State_ScreenUpdate: bool
'       record the Application.ScreenUpdating state before adjusting
'   State_DisplayStatusBar: bool
'       record the Application.DisplayStatusBar state before adjusting
'   State_EnableEvents: bool
'       record the Application.EnableEvents state before adjusting
'   State_CalcMode: xlCalculation
'       record the Application.Calculation state before adjusting
'   StatesSaved:    bool
'       track if update states were recorded
'#######################################################################################

Option Explicit

Public Enum UpdateState
    TurnOff
    TurnOn
End Enum

Dim State_ScreenUpdate As Boolean
Dim State_DisplayStatusBar As Boolean
Dim State_EnableEvents As Boolean
Dim State_CalcMode As XlCalculation
Dim StatesSaved As Boolean

Public Sub App_UpdateStatus(State As UpdateState, Optional UseStatusBar As Boolean = False)
'#######################################################################################
' App_UpdateStatus
'   Adjust application level settings to speed up code execution
'   TurnOff will save the current application states to be restored with TurnOn
'   If no states were recorded earlier during runtime then TurnOn will set all to active
'
' Args:
'   State: UpdateState (enum)
'       enum value of TurnOn/TurnOff
'   UseStatusBar: boolean (Default: False)
'       Use True if the macro is making use of the status bar, otherwise the status
'       messages will not be displayed
'
' Affects:
'   All module level variables and Excel Application level settings
'#######################################################################################

Select Case update_status
    Case TurnOff
        If Not StatesSaved Then
            State_ScreenUpdate = Application.ScreenUpdating
            State_DisplayStatusBar = Application.DisplayStatusBar
            State_EnableEvents = Application.EnableEvents
            State_CalcMode = Application.Calculation
            StatesSaved = True
        End If
        
        Application.ScreenUpdating = False
        Application.DisplayStatusBar = UseStatusBar
        Application.Calculation = xlCalculationManual
        Application.EnableEvents = False
    
    Case TurnOn
        If StatesSaved Then
            Application.ScreenUpdating = State_ScreenUpdate
            Application.DisplayStatusBar = State_DisplayStatusBar
            Application.EnableEvents = State_EnableEvents
            Application.Calculation = State_CalcMode
        Else
            Debug.Print "RestoreStates was called without initially saving Application update states. Applying default of all updates on."
            Application.ScreenUpdating = True
            Application.DisplayStatusBar = True
            Application.Calculation = xlCalculationAutomatic
            Application.EnableEvents = True
        End If
               
End Select

End Sub
