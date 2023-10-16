Attribute VB_Name = "f_pM_TemplatesCoreCompact"
' CORE, do not change
'============================================================================================
'   NAME:     f_pM_TemplatesCoreCompact
'============================================================================================
'   Purpose:  contains declaration and procedure templates
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
' 1.2.0    16.10.2023    gueleh    Initially created
'--------------------------------------------------------------------------------------------
'   BACKLOG
'   ''''''''''''''''''''
'   none
'============================================================================================
Option Explicit
Option Private Module

Private Const s_m_COMPONENT_NAME As String = "f_pM_TemplatesCoreCompact"

' Purpose: template for an entry level sub, compact version
Public Sub f_p_TemplateSubEntryLevelCompact()
'Fixed, don't change
   Dim oC_Me As New f_C_CallParams: oC_Me.s_prop_rw_ComponentName = s_m_COMPONENT_NAME
'>>>>>>> Your custom settings here
   f_p_StartProcessing e_f_p_ProcessingMode_AutoCalcOffOnSceenUpdatingOffOn
   With oC_Me
      .s_prop_rw_ProcedureName = "f_p_TemplateSubEntryLevelCompact" 'Name of the sub
      .b_prop_rw_SilentError = False 'False will display a message box - you should only do this on entry level
      .s_prop_rw_ErrorMessage = "Your message here." 'A message that properly informs the user and the devs (silent errors will be logged nonetheless)
      .SetCallArgs "No args" 'If the sub takes args put the here like ("sExample:=" & sExample, "lExample:=" & lExample)
   End With
'Fixed, don't change
   If oC_f_p_FrameworkSettings.b_prop_rw_ThisIsATestRun Then f_p_RegisterUnitTest oC_Me
Try: On Error GoTo Catch


'>>>>>>> Your code here
   
   
   
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

' Purpose: template for a non-trivial lower level procedure with error handling and execution control, compact version
Public Function b_f_p_TemplateLowerLevelCompact() As Boolean
'Fixed, don't change
   Dim oC_Me As New f_C_CallParams: oC_Me.s_prop_rw_ComponentName = s_m_COMPONENT_NAME: If oC_f_p_FrameworkSettings.b_prop_rw_ThisIsATestRun Then f_p_RegisterUnitTest oC_Me
'>>>>>>> Your custom settings here
   With oC_Me
      .s_prop_rw_ProcedureName = "b_f_p_TemplateLowerLevelCompact" 'Name of the function
      .b_prop_rw_SilentError = True 'False will display a message box - you should only do this on entry level
      .s_prop_rw_ErrorMessage = "Your message here." 'A message that properly informs the user and the devs (silent errors will be logged nonetheless)
      .SetCallArgs "No args" 'If the sub takes args put the here like ("sExample:=" & sExample, "lExample:=" & lExample)
   End With
'Fixed, don't change
Try: On Error GoTo Catch


'>>>>>>> Your code here



'End of your code <<<<<<<


'Fixed, don't change
Finally: On Error Resume Next


'>>>>>>> Your code here



'End of your code <<<<<<<


'>>>>>>> Your custom settings here
   If oC_Me.oC_prop_r_Error Is Nothing Then b_f_p_TemplateLowerLevelCompact = True 'reports execution as successful to caller
'Fixed, don't change
   Exit Function
HandleError: af_pM_ErrorHandling.af_p_Hook_ErrorHandling_LowerLevel


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
End Function


