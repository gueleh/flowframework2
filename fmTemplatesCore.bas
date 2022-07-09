Attribute VB_Name = "fmTemplatesCore"
' -------------------------------------------------------------------------------------------
' CORE, do not change
'============================================================================================
'   NAME:     fmTemplatesCore
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

Private Const msCOMPONENT_NAME As String = "fmTemplatesCore"

' Purpose: template for an entry level sub
' 0.1.0    20220709    gueleh    Initially created
Public Sub fTemplateSubEntryLevel()

'Fixed, don't change
   Dim clsMe As New fclsCallParams
   clsMe.sComponentName = msCOMPONENT_NAME
   
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Your custom settings here
   'Consult the manual for learning about the options for fStartProcessing
   fStartProcessing 'calling without args only inits the globals
   With clsMe
      .sProcedureName = "fTemplateSubEntryLevel" 'Name of the sub
      .bSilentError = False 'False will display a message box - you should only do this on entry level
      .sErrorMessage = "Your message here." 'A message that properly informs the user and the devs (silent errors will be logged nonetheless)
      .SetCallArgs "No args" 'If the sub takes args put the here like ("sExample:=" & sExample, "lExample:=" & lExample)
   End With
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

'Fixed, don't change
   If fgclsFrameworkSettings.bThisIsATestRun Then fRegisterUnitTest clsMe
Try:
   On Error GoTo Catch
   
'>>>>>>> Your code here
'TODO: Write fTemplateSubEntryLevel
      
   'Example for lower level call involving error handler (should always be the case for non-trivial procedures)
   'The intendation below is supposed to make it easier to discern these calls from other if blocks
   'The params of Err.Raise are the default for execution errors, you may define your own cases in module afmErrorHandling
      If Not _
   fbTemplateLowerLevel() _
      Then Err.Raise _
         fHandledErrorExecutionOfLowerLevelFunction, , _
         fsHandledErrorDescription(fHandledErrorExecutionOfLowerLevelFunction)
   
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
   If clsMe.clsError Is Nothing _
   Then fRegisterError clsMe, Err.Number, Err.Description
   If fgclsFrameworkSettings.bThisIsATestRun Then fRegisterExecutionError clsMe
   If fgclsFrameworkSettings.bDebugModeIsOn And Not clsMe.bResumedOnce Then
      clsMe.bResumedOnce = True: Stop: Resume
   Else
      fHandleError clsMe: Resume Finally
   End If
End Sub

' Purpose: template for a non-trivial lower level procedure with error handling and execution control
' 0.1.0    20220709    gueleh    Initially created
' Usage: if you need to return one or more values then declare these as ByRef args as in the template below, e.g. ByRef sReturnValue As String
Public Function fbTemplateLowerLevel() As Boolean

'Fixed, don't change
   Dim clsMe As New fclsCallParams
   clsMe.sComponentName = msCOMPONENT_NAME
   If fgclsFrameworkSettings.bThisIsATestRun Then fRegisterUnitTest clsMe

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Your custom settings here
   With clsMe
      .sProcedureName = "fbTemplateLowerLevel" 'Name of the function
      .bSilentError = True 'False will display a message box - you should only do this on entry level
      .sErrorMessage = "Your message here." 'A message that properly informs the user and the devs (silent errors will be logged nonetheless)
      .SetCallArgs "No args" 'If the sub takes args put the here like ("sExample:=" & sExample, "lExample:=" & lExample)
   End With
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

'Fixed, don't change
Try:
   On Error GoTo Catch

'>>>>>>> Your code here
'TODO: Write fbTemplateLowerLevel
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
   If clsMe.clsError Is Nothing Then fbTemplateLowerLevel = True 'reports execution as successful to caller
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

'Fixed, don't change
   Exit Function
Catch:
   If clsMe.clsError Is Nothing _
   Then fRegisterError clsMe, Err.Number, Err.Description
   If fgclsFrameworkSettings.bThisIsATestRun Then fRegisterExecutionError clsMe
   If fgclsFrameworkSettings.bDebugModeIsOn And Not clsMe.bResumedOnce Then
      clsMe.bResumedOnce = True: Stop: Resume
   Else
      fHandleError clsMe: Resume Finally
   End If

End Function
