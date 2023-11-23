Attribute VB_Name = "f_pM_EntryLevel"
' CORE, do not change
'============================================================================================
'   NAME:     f_pM_EntryLevel
'============================================================================================
'   Purpose:  entry level procedures which are part of the framework core
'   Access:   Private
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
' 0.10.0    04.08.2022    gueleh    Changed property names to meet new convention
'   0.9.0    03.08.2022    gueleh    Initially created, added f_p_EnterDevelopmentMode
'                                      Added f_p_LeaveDevelopmentMode
'--------------------------------------------------------------------------------------------
'   BACKLOG
'   ''''''''''''''''''''
'   none
'============================================================================================
Option Explicit
Option Private Module

Private Const s_m_COMPONENT_NAME As String = "f_pM_EntryLevel"


Public Sub f_p_ToggleDevelopmentMode()

   Dim oC_Me As New f_C_CallParams
   oC_Me.s_prop_rw_ComponentName = s_m_COMPONENT_NAME
   
   f_p_StartProcessing 'calling without args only inits the globals
   With oC_Me
      .s_prop_rw_ProcedureName = "f_p_ToggleDevelopmentMode" 'Name of the sub
      .b_prop_rw_SilentError = False 'False will display a message box - you should only do this on entry level
      .s_prop_rw_ErrorMessage = "Entering the development mode failed." 'A message that properly informs the user and the devs (silent errors will be logged nonetheless)
      .SetCallArgs "No args" 'If the sub takes args put the here like ("sExample:=" & sExample, "lExample:=" & lExample)
   End With

   If oC_f_p_FrameworkSettings.b_prop_rw_ThisIsATestRun Then f_p_RegisterUnitTest oC_Me

Try:
   On Error GoTo Catch
   Dim oC As New af_C_AppModes
   Dim sPassword As String
   
   If Not oC_f_p_FrameworkSettings.b_prop_r_DevelopmentModeIsOn Then
      sPassword = InputBox("Please enter development password")
      
      If Not oC.bPasswordDevModeCorrect(sPassword) Then
         MsgBox "The password is incorrect. Action aborted.", vbCritical
         GoTo Finally
      End If
   End If
   
      If Not _
   oC.bSetDevelopmentModeTo(Not oC_f_p_FrameworkSettings.b_prop_r_DevelopmentModeIsOn) _
      Then Err.Raise _
         e_f_p_HandledError_ExecutionOfLowerLevelFunction, , _
         s_f_p_HandledErrorDescription(e_f_p_HandledError_ExecutionOfLowerLevelFunction)
   
Finally:
   On Error Resume Next
   a_wks_Administration.Activate
   f_p_EndProcessing 'calling without args does nothing
   Exit Sub
   
Catch:
   If oC_Me.oC_prop_r_Error Is Nothing _
   Then f_p_RegisterError oC_Me, Err.Number, Err.Description
   If oC_f_p_FrameworkSettings.b_prop_rw_ThisIsATestRun Then f_p_RegisterExecutionError oC_Me
   If oC_f_p_FrameworkSettings.b_prop_r_DebugModeIsOn And Not oC_Me.b_prop_rw_ResumedOnce Then
      oC_Me.b_prop_rw_ResumedOnce = True: Stop: Resume
   Else
      f_p_HandleError oC_Me: Resume Finally
   End If
End Sub


' Purpose: enters the development mode, showing all wks, showing all names
' 0.9.0    03.08.2022    gueleh    Initially created
Public Sub f_p_EnterDevelopmentMode()

   Dim oC_Me As New f_C_CallParams
   oC_Me.s_prop_rw_ComponentName = s_m_COMPONENT_NAME
   
   f_p_StartProcessing 'calling without args only inits the globals
   With oC_Me
      .s_prop_rw_ProcedureName = "f_p_EnterDevelopmentMode" 'Name of the sub
      .b_prop_rw_SilentError = False 'False will display a message box - you should only do this on entry level
      .s_prop_rw_ErrorMessage = "Entering the development mode failed." 'A message that properly informs the user and the devs (silent errors will be logged nonetheless)
      .SetCallArgs "No args" 'If the sub takes args put the here like ("sExample:=" & sExample, "lExample:=" & lExample)
   End With

   If oC_f_p_FrameworkSettings.b_prop_rw_ThisIsATestRun Then f_p_RegisterUnitTest oC_Me

Try:
   On Error GoTo Catch
   Dim oC As New af_C_AppModes
   Dim sPassword As String
   
   sPassword = InputBox("Please enter development password")
   
   If Not oC.bPasswordDevModeCorrect(sPassword) Then
      MsgBox "The password is incorrect. Action aborted.", vbCritical
      GoTo Finally
   End If
      
      If Not _
   oC.bSetDevelopmentModeTo(True) _
      Then Err.Raise _
         e_f_p_HandledError_ExecutionOfLowerLevelFunction, , _
         s_f_p_HandledErrorDescription(e_f_p_HandledError_ExecutionOfLowerLevelFunction)
   
Finally:
   On Error Resume Next
   f_p_EndProcessing 'calling without args does nothing
   Exit Sub
   
Catch:
   If oC_Me.oC_prop_r_Error Is Nothing _
   Then f_p_RegisterError oC_Me, Err.Number, Err.Description
   If oC_f_p_FrameworkSettings.b_prop_rw_ThisIsATestRun Then f_p_RegisterExecutionError oC_Me
   If oC_f_p_FrameworkSettings.b_prop_r_DebugModeIsOn And Not oC_Me.b_prop_rw_ResumedOnce Then
      oC_Me.b_prop_rw_ResumedOnce = True: Stop: Resume
   Else
      f_p_HandleError oC_Me: Resume Finally
   End If
End Sub

' Purpose: leaves the development mode, setting all tech wks to very hidden, hiding all tech names, leaving debug mode
' 0.9.0    03.08.2022    gueleh    Initially created
Public Sub f_p_LeaveDevelopmentMode()

   Dim oC_Me As New f_C_CallParams
   oC_Me.s_prop_rw_ComponentName = s_m_COMPONENT_NAME
   
   f_p_StartProcessing 'calling without args only inits the globals
   With oC_Me
      .s_prop_rw_ProcedureName = "f_p_LeaveDevelopmentMode" 'Name of the sub
      .b_prop_rw_SilentError = False 'False will display a message box - you should only do this on entry level
      .s_prop_rw_ErrorMessage = "Leaving the development mode failed." 'A message that properly informs the user and the devs (silent errors will be logged nonetheless)
      .SetCallArgs "No args" 'If the sub takes args put the here like ("sExample:=" & sExample, "lExample:=" & lExample)
   End With

   If oC_f_p_FrameworkSettings.b_prop_rw_ThisIsATestRun Then f_p_RegisterUnitTest oC_Me

Try:
   On Error GoTo Catch
   Dim oC As New af_C_AppModes

      If Not _
   oC.bSetDevelopmentModeTo(False) _
      Then Err.Raise _
         e_f_p_HandledError_ExecutionOfLowerLevelFunction, , _
         s_f_p_HandledErrorDescription(e_f_p_HandledError_ExecutionOfLowerLevelFunction)
   
Finally:
   On Error Resume Next
   f_p_EndProcessing 'calling without args does nothing
   Exit Sub
   
Catch:
   If oC_Me.oC_prop_r_Error Is Nothing _
   Then f_p_RegisterError oC_Me, Err.Number, Err.Description
   If oC_f_p_FrameworkSettings.b_prop_rw_ThisIsATestRun Then f_p_RegisterExecutionError oC_Me
   If oC_f_p_FrameworkSettings.b_prop_r_DebugModeIsOn And Not oC_Me.b_prop_rw_ResumedOnce Then
      oC_Me.b_prop_rw_ResumedOnce = True: Stop: Resume
   Else
      f_p_HandleError oC_Me: Resume Finally
   End If
End Sub


