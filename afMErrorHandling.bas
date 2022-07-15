Attribute VB_Name = "afmErrorHandling"
' -------------------------------------------------------------------------------------------
' APP-SPECIFIC CORE MODULE - you have to migrate app contents manually in case of a template update
'============================================================================================
'   NAME:     afmErrorHandling
'============================================================================================
'   Purpose:  application-specific error handling
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

Private Const msCOMPONENT_NAME As String = "afmErrorHandling"

' the app-specific error cases
Public Enum afHandledErrors
   afHandledErrorGeneralError = 19999
'>>>>>>> Add your error enum cases here

'<<<<<<<
End Enum

' Purpose: returns the error description based on the provided error number
' Template Versions:
' 0.1.0    20220709    gueleh    Initially created
Public Function afsHandledErrorDescription(ByVal afeHandledError As afHandledErrors) As String
   Dim sDesc As String
   Select Case afeHandledError
      Case afHandledErrorGeneralError
         sDesc = "The app-specific error was not further specified."
'>>>>>>> Add your error description cases here
      'Case afHandledErrorYourValue
         'sDesc = "Your text."
'<<<<<<<
      Case Else
         sDesc = "No description defined for this error. You can do this in Function afsHandledErrorDescription in module afmErrorHandling."
   End Select
   afsHandledErrorDescription = sDesc
End Function

