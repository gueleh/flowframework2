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

Private Const smCOMPONENT_NAME As String = "fpMTemplatesCore"

' Purpose: template for an entry level sub
' 0.1.0    20220709    gueleh    Initially created
Public Sub fTemplateSubEntryLevel()

'Fixed, don't change
   Dim oCMe As New fCCallParams
   oCMe.sComponentName = smCOMPONENT_NAME
   
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Your custom settings here
   'Consult the manual for learning about the options for fStartProcessing
   fStartProcessing 'calling without args only inits the globals
   With oCMe
      .sProcedureName = "fTemplateSubEntryLevel" 'Name of the sub
      .bSilentError = False 'False will display a message box - you should only do this on entry level
      .sErrorMessage = "Your message here." 'A message that properly informs the user and the devs (silent errors will be logged nonetheless)
      .SetCallArgs "No args" 'If the sub takes args put the here like ("sExample:=" & sExample, "lExample:=" & lExample)
   End With
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

'Fixed, don't change
   If oCfgFrameworkSettings.bThisIsATestRun Then fRegisterUnitTest oCMe
Try:
   On Error GoTo Catch
   
'>>>>>>> Your code here
'TODO: Write fTemplateSubEntryLevel
      
   'Example for lower level call involving error handler (should always be the case for non-trivial procedures)
   'The intendation below is supposed to make it easier to discern these calls from other if blocks
   'The params of Err.Raise are the default for execution errors, you may define your own cases in module afmErrorHandling
      If Not _
   bfTemplateLowerLevel() _
      Then Err.Raise _
         efHandledErrorExecutionOfLowerLevelFunction, , _
         sfHandledErrorDescription(efHandledErrorExecutionOfLowerLevelFunction)
   
'End of your code <<<<<<<
   
'Fixed, don't change
Finally:
   On Error Resume Next

'>>>>>>> Your code here
   'everything that must be executed regardless of an error or not
'End of your code <<<<<<<

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Your custom settings here
   'Consult the manual for learning about the options for fEndProcessing
   fEndProcessing 'calling without args does nothing
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

'Fixed, don't change
   Exit Sub
Catch:
   If oCMe.oCError Is Nothing _
   Then fRegisterError oCMe, Err.Number, Err.Description
   If oCfgFrameworkSettings.bThisIsATestRun Then fRegisterExecutionError oCMe
   If oCfgFrameworkSettings.bDebugModeIsOn And Not oCMe.bResumedOnce Then
      oCMe.bResumedOnce = True: Stop: Resume
   Else
      fHandleError oCMe: Resume Finally
   End If
End Sub

' Purpose: template for a non-trivial lower level procedure with error handling and execution control
' 0.1.0    20220709    gueleh    Initially created
' Usage: if you need to return one or more values then declare these as ByRef args as in the template below, e.g. ByRef sReturnValue As String
Public Function bfTemplateLowerLevel() As Boolean

'Fixed, don't change
   Dim oCMe As New fCCallParams
   oCMe.sComponentName = smCOMPONENT_NAME
   If oCfgFrameworkSettings.bThisIsATestRun Then fRegisterUnitTest oCMe

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Your custom settings here
   With oCMe
      .sProcedureName = "bfTemplateLowerLevel" 'Name of the function
      .bSilentError = True 'False will display a message box - you should only do this on entry level
      .sErrorMessage = "Your message here." 'A message that properly informs the user and the devs (silent errors will be logged nonetheless)
      .SetCallArgs "No args" 'If the sub takes args put the here like ("sExample:=" & sExample, "lExample:=" & lExample)
   End With
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

'Fixed, don't change
Try:
   On Error GoTo Catch

'>>>>>>> Your code here
'TODO: Write bfTemplateLowerLevel
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
   If oCMe.oCError Is Nothing Then bfTemplateLowerLevel = True 'reports execution as successful to caller
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

'Fixed, don't change
   Exit Function
Catch:
   If oCMe.oCError Is Nothing _
   Then fRegisterError oCMe, Err.Number, Err.Description
   If oCfgFrameworkSettings.bThisIsATestRun Then fRegisterExecutionError oCMe
   If oCfgFrameworkSettings.bDebugModeIsOn And Not oCMe.bResumedOnce Then
      oCMe.bResumedOnce = True: Stop: Resume
   Else
      fHandleError oCMe: Resume Finally
   End If

End Function
