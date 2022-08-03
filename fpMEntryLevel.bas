Attribute VB_Name = "fpMEntryLevel"
' -------------------------------------------------------------------------------------------
' CORE, do not change
'============================================================================================
'   NAME:     fpMEntryLevel
'============================================================================================
'   Purpose:  entry level procedures which are part of the framework core
'   Access:   Public
'   Type:     Module
'   Author:   Günther Lehner
'   Contact:  guenther.lehner@protonmail.com
'   GitHubID: gueleh
'   Required:
'   Usage:
'--------------------------------------------------------------------------------------------
'   VERSION HISTORY
'   Version    Date    Developer    Changes
'   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   0.9.0    03.08.2022    gueleh    Initially created, added f_p_EnterDevelopmentMode
'                                      Added f_p_LeaveDevelopmentMode
'--------------------------------------------------------------------------------------------
'   BACKLOG
'   ''''''''''''''''''''
'   none
'============================================================================================
Option Explicit
Option Private Module

Private Const s_m_COMPONENT_NAME As String = "fpMEntryLevel"

' Purpose: enters the development mode, showing all wks, showing all names
' 0.9.0    03.08.2022    gueleh    Initially created
Public Sub f_p_EnterDevelopmentMode()

   Dim oC_Me As New fCCallParams
   oC_Me.sComponentName = s_m_COMPONENT_NAME
   
   f_p_StartProcessing 'calling without args only inits the globals
   With oC_Me
      .sProcedureName = "f_p_EnterDevelopmentMode" 'Name of the sub
      .bSilentError = False 'False will display a message box - you should only do this on entry level
      .sErrorMessage = "Entering the development mode failed." 'A message that properly informs the user and the devs (silent errors will be logged nonetheless)
      .SetCallArgs "No args" 'If the sub takes args put the here like ("sExample:=" & sExample, "lExample:=" & lExample)
   End With

   If oC_f_p_FrameworkSettings.bThisIsATestRun Then f_p_RegisterUnitTest oC_Me

Try:
   On Error GoTo Catch
   
      If Not _
   b_f_p_SetDevelopmentModeTo(True) _
      Then Err.Raise _
         e_f_p_HandledError_ExecutionOfLowerLevelFunction, , _
         s_f_p_HandledErrorDescription(e_f_p_HandledError_ExecutionOfLowerLevelFunction)
   
Finally:
   On Error Resume Next
   f_p_EndProcessing 'calling without args does nothing
   Exit Sub
   
Catch:
   If oC_Me.oCError Is Nothing _
   Then f_p_RegisterError oC_Me, Err.Number, Err.Description
   If oC_f_p_FrameworkSettings.bThisIsATestRun Then f_p_RegisterExecutionError oC_Me
   If oC_f_p_FrameworkSettings.bDebugModeIsOn And Not oC_Me.bResumedOnce Then
      oC_Me.bResumedOnce = True: Stop: Resume
   Else
      f_p_HandleError oC_Me: Resume Finally
   End If
End Sub

' Purpose: leaves the development mode, setting all tech wks to very hidden, hiding all tech names, leaving debug mode
' 0.9.0    03.08.2022    gueleh    Initially created
Public Sub f_p_LeaveDevelopmentMode()

   Dim oC_Me As New fCCallParams
   oC_Me.sComponentName = s_m_COMPONENT_NAME
   
   f_p_StartProcessing 'calling without args only inits the globals
   With oC_Me
      .sProcedureName = "f_p_LeaveDevelopmentMode" 'Name of the sub
      .bSilentError = False 'False will display a message box - you should only do this on entry level
      .sErrorMessage = "Leaving the development mode failed." 'A message that properly informs the user and the devs (silent errors will be logged nonetheless)
      .SetCallArgs "No args" 'If the sub takes args put the here like ("sExample:=" & sExample, "lExample:=" & lExample)
   End With

   If oC_f_p_FrameworkSettings.bThisIsATestRun Then f_p_RegisterUnitTest oC_Me

Try:
   On Error GoTo Catch
   
      If Not _
   b_f_p_SetDevelopmentModeTo(False) _
      Then Err.Raise _
         e_f_p_HandledError_ExecutionOfLowerLevelFunction, , _
         s_f_p_HandledErrorDescription(e_f_p_HandledError_ExecutionOfLowerLevelFunction)
   
Finally:
   On Error Resume Next
   f_p_EndProcessing 'calling without args does nothing
   Exit Sub
   
Catch:
   If oC_Me.oCError Is Nothing _
   Then f_p_RegisterError oC_Me, Err.Number, Err.Description
   If oC_f_p_FrameworkSettings.bThisIsATestRun Then f_p_RegisterExecutionError oC_Me
   If oC_f_p_FrameworkSettings.bDebugModeIsOn And Not oC_Me.bResumedOnce Then
      oC_Me.bResumedOnce = True: Stop: Resume
   Else
      f_p_HandleError oC_Me: Resume Finally
   End If
End Sub


