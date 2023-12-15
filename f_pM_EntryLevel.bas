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

Public Sub f_p_ToggleMaintenanceMode()

   Dim oC_Me As New f_C_CallParams
   oC_Me.s_prop_rw_ComponentName = s_m_COMPONENT_NAME
   
   f_p_StartProcessing e_f_p_ProcessingMode_AutoCalcOffOnSceenUpdatingOffOn
   With oC_Me
      .s_prop_rw_ProcedureName = "f_p_ToggleMaintenanceMode" 'Name of the sub
      .b_prop_rw_SilentError = False 'False will display a message box - you should only do this on entry level
      .s_prop_rw_ErrorMessage = "Entering the development mode failed." 'A message that properly informs the user and the devs (silent errors will be logged nonetheless)
      .SetCallArgs "No args" 'If the sub takes args put the here like ("sExample:=" & sExample, "lExample:=" & lExample)
   End With

   If oC_f_p_FrameworkSettings.b_prop_rw_ThisIsATestRun Then f_p_RegisterUnitTest oC_Me

Try:
   On Error GoTo Catch
   Dim oC As New af_C_AppModes
   Dim sPassword As String
   
   If oC_f_p_FrameworkSettings.b_prop_r_DevelopmentModeIsOn Then
      MsgBox "Please leave development mode first and then try again.", vbInformation
      GoTo Finally
   End If
   
   If Not oC_f_p_FrameworkSettings.b_prop_r_MaintenanceModeIsOn Then
      sPassword = InputBox("Please enter admin password")
      
      If Not oC.bPasswordMaintenanceModeCorrect(sPassword) Then
         MsgBox "The password is incorrect. Action aborted.", vbCritical
         GoTo Finally
      End If
   End If
   
      If Not _
   oC.bSetMaintenanceModeTo(Not oC_f_p_FrameworkSettings.b_prop_r_MaintenanceModeIsOn) _
      Then Err.Raise _
         e_f_p_HandledError_ExecutionOfLowerLevelFunction, , _
         s_f_p_HandledErrorDescription(e_f_p_HandledError_ExecutionOfLowerLevelFunction)
   
Finally:
   On Error Resume Next
   If oC_f_p_FrameworkSettings.b_prop_r_MaintenanceModeIsOn Then
      a_wks_Administration.Activate
   Else
      a_wks_Main.Activate
   End If
   f_p_EndProcessing e_f_p_ProcessingMode_AutoCalcOffOnSceenUpdatingOffOn
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


Public Sub f_p_ToggleDevelopmentMode()

   Dim oC_Me As New f_C_CallParams
   oC_Me.s_prop_rw_ComponentName = s_m_COMPONENT_NAME
   
   f_p_StartProcessing e_f_p_ProcessingMode_AutoCalcOffOnSceenUpdatingOffOn
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
   
   If Not oC_f_p_FrameworkSettings.b_prop_r_MaintenanceModeIsOn Then
      MsgBox "Please enter maintenance mode first and then try again.", vbInformation
      GoTo Finally
   End If
   
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
   f_p_EndProcessing e_f_p_ProcessingMode_AutoCalcOffOnSceenUpdatingOffOn
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

Public Sub f_p_DeployWorkbook()
'Fixed, don't change
   Dim oC_Me As New f_C_CallParams: oC_Me.s_prop_rw_ComponentName = s_m_COMPONENT_NAME
'>>>>>>> Your custom settings here
   f_p_StartProcessing e_f_p_ProcessingMode_AutoCalcOffOnSceenUpdatingOffOn
   With oC_Me
      .s_prop_rw_ProcedureName = "f_p_DeployWorkbook" 'Name of the sub
      .b_prop_rw_SilentError = False 'False will display a message box - you should only do this on entry level
      .s_prop_rw_ErrorMessage = "Deployment failed." 'A message that properly informs the user and the devs (silent errors will be logged nonetheless)
      .SetCallArgs "No args" 'If the sub takes args put the here like ("sExample:=" & sExample, "lExample:=" & lExample)
   End With
'Fixed, don't change
   If oC_f_p_FrameworkSettings.b_prop_rw_ThisIsATestRun Then f_p_RegisterUnitTest oC_Me
Try: On Error GoTo Catch
'>>>>>>> Your code here
   If Not MsgBox("This will save a copy of this workbook and remove all code components which are only needed in development. Continue?", vbYesNo) = vbYes Then
      MsgBox "Deployment aborted.", vbInformation
      GoTo Finally
   End If
   
   Dim oC As New f_C_Deploy
   
      If Not _
   oC.bSaveAsProdAndRemoveDEVModules() _
      Then Err.Raise _
         e_f_p_HandledError_ExecutionOfLowerLevelFunction, , _
         s_f_p_HandledErrorDescription(e_f_p_HandledError_ExecutionOfLowerLevelFunction)
   
   MsgBox "Copy without DEV code saved. Leaving dev and maintenance mode has to be done separately, if applicable.", vbOKOnly
'End of your code <<<<<<<
   

'Fixed, don't change
Finally: On Error Resume Next


'>>>>>>> Your code here



'End of your code <<<<<<<


'>>>>>>> Your custom settings here
   f_p_EndProcessing e_f_p_ProcessingMode_AutoCalcOffOnSceenUpdatingOffOn
'Fixed, don't change
   Exit Sub
HandleError: af_pM_ErrorHandling.af_p_Hook_ErrorHandling_EntryLevel


'>>>>>>> Your code here



'End of your code <<<<<<<
   
   
'Fixed, don't change
   Resume Finally
Catch:
   If oC_Me.oC_prop_r_Error Is Nothing Then f_p_RegisterError oC_Me, Err.Number, Err.Description
   If oC_f_p_FrameworkSettings.b_prop_rw_ThisIsATestRun Then f_p_RegisterExecutionError oC_Me
   If oC_f_p_FrameworkSettings.b_prop_r_DebugModeIsOn And Not oC_Me.b_prop_rw_ResumedOnce Then
      oC_Me.b_prop_rw_ResumedOnce = True: Stop: Resume
   Else
      f_p_HandleError oC_Me: GoTo HandleError
   End If
End Sub


