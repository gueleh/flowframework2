Attribute VB_Name = "afpMErrorHandling"
' -------------------------------------------------------------------------------------------
' APP-SPECIFIC CORE MODULE - you have to migrate app contents manually in case of a template update
'============================================================================================
'   NAME:     afpMErrorHandling
'============================================================================================
'   Purpose:  application-specific error handling
'   Access:   Private
'   Type:     Module
'   Author:   G�nther Lehner
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

Private Const s_m_COMPONENT_NAME As String = "afpMErrorHandling"

' the app-specific error cases
Public Enum e_af_p_HandledErrors
   e_af_p_HandledError_GeneralError = 19999
'>>>>>>> Add your error enum cases here - add cases for them to safHandledErrorDescription below
' if you want to have specific error descriptions

'<<<<<<<
End Enum

' Purpose: returns the error description based on the provided error number
' Template Versions:
' 0.1.0    20220709    gueleh    Initially created
Public Function s_af_p_HandledErrorDescription(ByVal eafHandledError As e_af_p_HandledErrors) As String
   Dim sDesc As String
   Select Case eafHandledError
      Case e_af_p_HandledError_GeneralError
         sDesc = "The app-specific error was not further specified."
'>>>>>>> Add your error description cases here
      'Case e_af_p_HandledError_YourValue
         'sDesc = "Your text."
'<<<<<<<
      Case Else
         sDesc = "No description defined for this error. You can do this in Function af_s_p_HandledErrorDescription in module afpMErrorHandling."
   End Select
   s_af_p_HandledErrorDescription = sDesc
End Function


