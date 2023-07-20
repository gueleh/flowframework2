Attribute VB_Name = "f_pM_TemplatesCore"
' CORE, do not change
'============================================================================================
'   NAME:     f_pM_TemplatesCore
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
' 0.13.0    31.01.2023    gueleh    Added "ping" as indicator for execution in skeletons in both templates
' 0.10.0    04.08.2022    gueleh    Changed property names to meet new convention
'   0.1.0    20220709    gueleh    Initially created
'--------------------------------------------------------------------------------------------
'   BACKLOG
'   ''''''''''''''''''''
'   none
'============================================================================================
Option Explicit
Option Private Module

Private Const s_m_COMPONENT_NAME As String = "f_pM_TemplatesCore"

' Purpose: template for an entry level sub
' 0.13.0    31.01.2023    gueleh    Added "ping" as indicator for execution in skeletons
' 0.1.0    20220709    gueleh    Initially created
Public Sub f_p_TemplateSubEntryLevel()

'Fixed, don't change
   Dim oC_Me As New f_C_CallParams
   oC_Me.s_prop_rw_ComponentName = s_m_COMPONENT_NAME
   
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Your custom settings here
   'Consult the manual for learning about the options for f_p_StartProcessing
   ' Default: init globals, turning off and on screen updating and automatic calculation
   f_p_StartProcessing e_f_p_ProcessingMode_AutoCalcOffOnSceenUpdatingOffOn
   
   ' Alternative: only initializing globals
   'f_p_StartProcessing 'calling without args only inits the globals
   
   ' Alternative: app specific behavior, determined by own enum values in the second argument
   'f_p_StartProcessing e_f_p_ProcessingMode_AppSpecific, e_af_p_ProcessingModeGlobalsOnly
   
   With oC_Me
      .s_prop_rw_ProcedureName = "f_p_TemplateSubEntryLevel" 'Name of the sub
      .b_prop_rw_SilentError = False 'False will display a message box - you should only do this on entry level
      .s_prop_rw_ErrorMessage = "Your message here." 'A message that properly informs the user and the devs (silent errors will be logged nonetheless)
      .SetCallArgs "No args" 'If the sub takes args put the here like ("sExample:=" & sExample, "lExample:=" & lExample)
   End With
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

'Fixed, don't change
   If oC_f_p_FrameworkSettings.b_prop_rw_ThisIsATestRun Then f_p_RegisterUnitTest oC_Me
Try:
   On Error GoTo Catch
   
'>>>>>>> Your code here
'TODO: Write f_g_TemplateSubEntryLevel
      
'TODO: Remove f_p_PrintCallParams oC_Me if not needed - it just shows via printing to the direct window if the proc is executed
   f_p_PrintCallParams oC_Me
      
   'Example for lower level call involving error handler (should always be the case for non-trivial procedures)
   'The intendation below is supposed to make it easier to discern these calls from other if blocks
   'The params of Err.Raise are the default for execution errors, you may define your own cases in module afmErrorHandling
      If Not _
   b_f_p_TemplateLowerLevel() _
      Then Err.Raise _
         e_f_p_HandledError_ExecutionOfLowerLevelFunction, , _
         s_f_p_HandledErrorDescription(e_f_p_HandledError_ExecutionOfLowerLevelFunction)
   
'End of your code <<<<<<<
   
'Fixed, don't change
Finally:
   On Error Resume Next

'>>>>>>> Your code here
   'everything that must be executed regardless of an error or not
'End of your code <<<<<<<

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Your custom settings here
   'Consult the manual for learning about the options for f_p_EndProcessing
   f_p_EndProcessing e_f_p_ProcessingMode_AutoCalcOffOnSceenUpdatingOffOn
   'Alternative:
   'f_p_EndProcessing 'calling without args does nothing
   'Alternative:
   'f_p_EndProcessing e_f_p_ProcessingMode_AppSpecific, e_af_p_ProcessingModeGlobalsOnly
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

'Fixed, don't change
   Exit Sub
   
HandleError:
   af_pM_ErrorHandling.af_p_Hook_ErrorHandling_EntryLevel
'>>>>>>> Your code here
   'everything that must be executed in case of an error

'End of your code <<<<<<<
   Resume Finally

'Fixed, don't change
Catch:
   If oC_Me.oC_prop_r_Error Is Nothing _
   Then f_p_RegisterError oC_Me, Err.Number, Err.Description
   If oC_f_p_FrameworkSettings.b_prop_rw_ThisIsATestRun Then f_p_RegisterExecutionError oC_Me
   If oC_f_p_FrameworkSettings.b_prop_r_DebugModeIsOn And Not oC_Me.b_prop_rw_ResumedOnce Then
      oC_Me.b_prop_rw_ResumedOnce = True: Stop: Resume
   Else
      f_p_HandleError oC_Me: GoTo HandleError
   End If
End Sub

' Purpose: template for a non-trivial lower level procedure with error handling and execution control
' 0.13.0    31.01.2023    gueleh    Added "ping" as indicator for execution in skeletons
' 0.1.0    20220709    gueleh    Initially created
' Usage: if you need to return one or more values then declare these as ByRef args as in the template below, e.g. ByRef sReturnValue As String
Public Function b_f_p_TemplateLowerLevel() As Boolean

'Fixed, don't change
   Dim oC_Me As New f_C_CallParams
   oC_Me.s_prop_rw_ComponentName = s_m_COMPONENT_NAME
   If oC_f_p_FrameworkSettings.b_prop_rw_ThisIsATestRun Then f_p_RegisterUnitTest oC_Me

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Your custom settings here
   With oC_Me
      .s_prop_rw_ProcedureName = "b_f_p_TemplateLowerLevel" 'Name of the function
      .b_prop_rw_SilentError = True 'False will display a message box - you should only do this on entry level
      .s_prop_rw_ErrorMessage = "Your message here." 'A message that properly informs the user and the devs (silent errors will be logged nonetheless)
      .SetCallArgs "No args" 'If the sub takes args put the here like ("sExample:=" & sExample, "lExample:=" & lExample)
   End With
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

'Fixed, don't change
Try:
   On Error GoTo Catch

'>>>>>>> Your code here
'TODO: Write b_f_g_TemplateLowerLevel

'TODO: Remove f_p_PrintCallParams oC_Me if not needed - it just shows via printing to the direct window if the proc is executed
   f_p_PrintCallParams oC_Me
      
'End of your code <<<<<<<

'Fixed, don't change
Finally:
   On Error Resume Next

'>>>>>>> Your code here
   'everything that must be executed regardless of an error or not
'End of your code <<<<<<<

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Your custom settings here
'MAKE SURE TO REPLACE fbTemplateLowerLevel WITH THE NAME OF YOUR FUNCTION
   If oC_Me.oC_prop_r_Error Is Nothing Then b_f_p_TemplateLowerLevel = True 'reports execution as successful to caller
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

'Fixed, don't change
   Exit Function
   
HandleError:
   af_pM_ErrorHandling.af_p_Hook_ErrorHandling_LowerLevel
'>>>>>>> Your code here
   'everything that must be executed in case of an error

'End of your code <<<<<<<

'Fixed, don't change
   Resume Finally
Catch:
   If oC_Me.oC_prop_r_Error Is Nothing _
   Then f_p_RegisterError oC_Me, Err.Number, Err.Description
   If oC_f_p_FrameworkSettings.b_prop_rw_ThisIsATestRun Then f_p_RegisterExecutionError oC_Me
   If oC_f_p_FrameworkSettings.b_prop_r_DebugModeIsOn And Not oC_Me.b_prop_rw_ResumedOnce Then
      oC_Me.b_prop_rw_ResumedOnce = True: Stop: Resume
   Else
      f_p_HandleError oC_Me: GoTo HandleError
   End If

End Function
