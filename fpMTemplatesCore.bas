Attribute VB_Name = "fpMTemplatesCore"
' -------------------------------------------------------------------------------------------
' CORE, do not change
'============================================================================================
'   NAME:     fpMTemplatesCore
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
'   0.1.0    20220709    gueleh    Initially created
'--------------------------------------------------------------------------------------------
'   BACKLOG
'   ''''''''''''''''''''
'   none
'============================================================================================
Option Explicit
Option Private Module

Private Const s_m_COMPONENT_NAME As String = "fpMTemplatesCore"

' Purpose: template for an entry level sub
' 0.1.0    20220709    gueleh    Initially created
Public Sub f_g_TemplateSubEntryLevel()

'Fixed, don't change
   Dim oC_Me As New fCCallParams
   oC_Me.sComponentName = s_m_COMPONENT_NAME
   
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Your custom settings here
   'Consult the manual for learning about the options for f_g_StartProcessing
   f_g_StartProcessing 'calling without args only inits the globals
   With oC_Me
      .sProcedureName = "f_g_TemplateSubEntryLevel" 'Name of the sub
      .bSilentError = False 'False will display a message box - you should only do this on entry level
      .sErrorMessage = "Your message here." 'A message that properly informs the user and the devs (silent errors will be logged nonetheless)
      .SetCallArgs "No args" 'If the sub takes args put the here like ("sExample:=" & sExample, "lExample:=" & lExample)
   End With
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

'Fixed, don't change
   If oC_f_g_FrameworkSettings.bThisIsATestRun Then f_g_RegisterUnitTest oC_Me
Try:
   On Error GoTo Catch
   
'>>>>>>> Your code here
'TODO: Write f_g_TemplateSubEntryLevel
      
   'Example for lower level call involving error handler (should always be the case for non-trivial procedures)
   'The intendation below is supposed to make it easier to discern these calls from other if blocks
   'The params of Err.Raise are the default for execution errors, you may define your own cases in module afmErrorHandling
      If Not _
   b_f_g_TemplateLowerLevel() _
      Then Err.Raise _
         e_f_g_HandledError_ExecutionOfLowerLevelFunction, , _
         s_f_g_HandledErrorDescription(e_f_g_HandledError_ExecutionOfLowerLevelFunction)
   
'End of your code <<<<<<<
   
'Fixed, don't change
Finally:
   On Error Resume Next

'>>>>>>> Your code here
   'everything that must be executed regardless of an error or not
'End of your code <<<<<<<

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Your custom settings here
   'Consult the manual for learning about the options for f_g_EndProcessing
   f_g_EndProcessing 'calling without args does nothing
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

'Fixed, don't change
   Exit Sub
Catch:
   If oC_Me.oCError Is Nothing _
   Then f_g_RegisterError oC_Me, Err.Number, Err.Description
   If oC_f_g_FrameworkSettings.bThisIsATestRun Then f_g_RegisterExecutionError oC_Me
   If oC_f_g_FrameworkSettings.bDebugModeIsOn And Not oC_Me.bResumedOnce Then
      oC_Me.bResumedOnce = True: Stop: Resume
   Else
      f_g_HandleError oC_Me: Resume Finally
   End If
End Sub

' Purpose: template for a non-trivial lower level procedure with error handling and execution control
' 0.1.0    20220709    gueleh    Initially created
' Usage: if you need to return one or more values then declare these as ByRef args as in the template below, e.g. ByRef sReturnValue As String
Public Function b_f_g_TemplateLowerLevel() As Boolean

'Fixed, don't change
   Dim oC_Me As New fCCallParams
   oC_Me.sComponentName = s_m_COMPONENT_NAME
   If oC_f_g_FrameworkSettings.bThisIsATestRun Then f_g_RegisterUnitTest oC_Me

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Your custom settings here
   With oC_Me
      .sProcedureName = "b_f_g_TemplateLowerLevel" 'Name of the function
      .bSilentError = True 'False will display a message box - you should only do this on entry level
      .sErrorMessage = "Your message here." 'A message that properly informs the user and the devs (silent errors will be logged nonetheless)
      .SetCallArgs "No args" 'If the sub takes args put the here like ("sExample:=" & sExample, "lExample:=" & lExample)
   End With
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

'Fixed, don't change
Try:
   On Error GoTo Catch

'>>>>>>> Your code here
'TODO: Write b_f_g_TemplateLowerLevel
   Debug.Print 1 / 0
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
   If oC_Me.oCError Is Nothing Then b_f_g_TemplateLowerLevel = True 'reports execution as successful to caller
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

'Fixed, don't change
   Exit Function
Catch:
   If oC_Me.oCError Is Nothing _
   Then f_g_RegisterError oC_Me, Err.Number, Err.Description
   If oC_f_g_FrameworkSettings.bThisIsATestRun Then f_g_RegisterExecutionError oC_Me
   If oC_f_g_FrameworkSettings.bDebugModeIsOn And Not oC_Me.bResumedOnce Then
      oC_Me.bResumedOnce = True: Stop: Resume
   Else
      f_g_HandleError oC_Me: Resume Finally
   End If

End Function
