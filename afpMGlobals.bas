Attribute VB_Name = "afpMGlobals"
' -------------------------------------------------------------------------------------------
' APP-SPECIFIC CORE, your content has to be migrated manually if the template if updated
'============================================================================================
'   NAME:     afpMGlobals
'============================================================================================
'   Purpose:  the app-specific globals of the framework
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

Private Const smCOMPONENT_NAME As String = "afpMGlobals"

'determines which mode supposed to be executed
Public Enum eafProcessingModes
   eafProcessingModeGlobalsOnly
'>>>>>>> add your modes here

'<<<<<<<
End Enum

' Purpose: executes the start processing logic as determined by the app-specific case
' Template Versions:
' 0.1.0    20220709    gueleh    Initially created
Public Sub afStartProcessingMode(ByVal eafProcessingMode As eafProcessingModes)
   Select Case eafProcessingMode
      Case eafProcessingModeGlobalsOnly
         'Do nothing except for the required initialization
'>>>>>>> Your cases here
         
'<<<<<<<
      Case Else
         'Do nothing except for the required initialization
   End Select
End Sub

' Purpose: executes the start processing logic as determined by the app-specific case
' Template Versions:
' 0.1.0    20220709    gueleh    Initially created
Public Sub afEndProcessingMode(ByVal eafProcessingMode As eafProcessingModes)
   Select Case eafProcessingMode
      Case eafProcessingModeGlobalsOnly
         'Do nothing
'>>>>>>> Your cases here
         
'<<<<<<<
      Case Else
         'Do nothing
   End Select
End Sub

