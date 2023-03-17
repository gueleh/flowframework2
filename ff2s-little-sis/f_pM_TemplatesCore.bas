Attribute VB_Name = "f_pM_TemplatesCore"
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
' 0.1.0    17.03.2023    gueleh    Initially created
'--------------------------------------------------------------------------------------------
'   BACKLOG
'   ''''''''''''''''''''
'   none
'============================================================================================
Option Explicit
Option Private Module

Private Const s_m_COMPONENT_NAME As String = "fpMTemplatesCore"

' Purpose: template for an entry level sub
' 0.1.0    17.03.2023    gueleh    Initially created
Public Sub f_p_TemplateSubEntryLevel()

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'Change name of the sub if you want to have this information in the error log
   Const sNAME_OF_SUB As String = "f_p_TemplateSubEntryLevel"
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

'Fixed, don't change
   Dim bResumedOnce As Boolean
   Dim oCError As f_C_Error
   Dim bIsSilentError As Boolean
   Dim sErrorMessage As String
   
Try:
   On Error GoTo Catch
   af_p_StartProcessing 'turns off screen updating, unprotects sheets etc.
'End Fixed
   
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>> Your code here
   
   'Change this is necessary
   bIsSilentError = False 'if False, a MsgBox will be displayed. Otherwise only an entry in the error log.
   sErrorMessage = "The error message the user should see."

'TODO: Write f_p_TemplateSubEntryLevel
      
   'Example for lower level call involving error handler (should always be the case for non-trivial procedures)
   'The intendation below is supposed to make it easier to discern these calls from other if blocks
      If Not _
   b_f_p_TemplateLowerLevel() _
      Then Err.Raise l_f_p_EXECUTION_ERROR_NUMBER, , s_f_p_EXECUTION_ERROR_TEXT
   
'End of your code <<<<<<<
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
   
'Fixed, don't change
Finally:
   On Error Resume Next
'End Fixed

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>> Your code here
'>>>>>>> everything that MUST BE EXECUTED regardless of an error or not


'End of your code <<<<<<<
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

'Fixed, don't change
   af_p_EndProcessing 'turns on screen updating, protects sheets etc.
   Exit Sub

Catch:
   'Set error data for logging if not already existent
   If oCError Is Nothing Then
      Set oCError = New f_C_Error
      oCError.SetErrorData _
         Err.Number, Err.Description, sNAME_OF_SUB, _
         s_m_COMPONENT_NAME, bIsSilentError, sErrorMessage
   End If
   
   'If in debug mode, then the code will stop once so that you can step into
   '  the row which caused the error.
   '  If already stopped and resumed once, the error data are sent to the error handler
   If b_p_DEBUG_MODE_IS_ON _
   And Not bResumedOnce Then
      bResumedOnce = True: Stop: Resume
   Else
      f_p_LogError oCError: Resume Finally
   End If
'End Fixed

End Sub

' Purpose: template for a non-trivial lower level procedure with error handling and execution control
' 0.1.0    17.03.2023    gueleh    Initially created
' Usage: if you need to return one or more values then declare these as ByRef args as in the template below, e.g. ByRef sReturnValue As String
Public Function b_f_p_TemplateLowerLevel() As Boolean

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'Change name of the sub if you want to have this information in the error log
   Const sNAME_OF_FUNCTION As String = "b_f_p_TemplateLowerLevel"
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

'Fixed, don't change
   Dim bResumedOnce As Boolean
   Dim bExecutedSuccessfully As Boolean
   Dim oCError As f_C_Error
   Dim bIsSilentError As Boolean
   Dim sErrorMessage As String

Try:
   On Error GoTo Catch
   bExecutedSuccessfully = True
'End Fixed
   
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>> Your code here

   'Change this is necessary
   bIsSilentError = True 'if False, a MsgBox will be displayed. Otherwise only an entry in the error log.
   sErrorMessage = "The error message for the error log."

   Debug.Print 1 / 0 'to cause an error

'End of your code <<<<<<<
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


'Fixed, don't change
Finally:
   On Error Resume Next

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>> Your code here
'>>>>>>> everything that MUST BE EXECUTED regardless of an error or not
      
   

   'change this to meet the name of your function
   b_f_p_TemplateLowerLevel = bExecutedSuccessfully


'End of your code <<<<<<<
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


'Fixed, don't change
   Exit Function
Catch:
   'for informing the caller that the execution failed
   bExecutedSuccessfully = False
   
   'Set error data for logging if not already existent
   If oCError Is Nothing Then
      Set oCError = New f_C_Error
      oCError.SetErrorData _
         Err.Number, Err.Description, sNAME_OF_FUNCTION, _
         s_m_COMPONENT_NAME, bIsSilentError, sErrorMessage
   End If
   
   'If in debug mode, then the code will stop once so that you can step into
   '  the row which caused the error.
   '  If already stopped and resumed once, the error data are sent to the error handler
   If b_p_DEBUG_MODE_IS_ON _
   And Not bResumedOnce Then
      bResumedOnce = True: Stop: Resume
   Else
      f_p_LogError oCError: Resume Finally
   End If
'End Fixed

End Function
