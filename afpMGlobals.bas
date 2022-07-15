Attribute VB_Name = "afmGlobals"
' -------------------------------------------------------------------------------------------
' APP-SPECIFIC CORE, your content has to be migrated manually if the template if updated
'============================================================================================
'   NAME:     afmGlobals
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

Private Const msCOMPONENT_NAME As String = "afmGlobals"

'determines which mode supposed to be executed
Public Enum afProcessingModes
   afProcessingModeGlobalsOnly
'>>>>>>> add your modes here

'<<<<<<<
End Enum

' Purpose: executes the start processing logic as determined by the app-specific case
' Template Versions:
' 0.1.0    20220709    gueleh    Initially created
Public Sub afStartProcessingMode(ByVal afeProcessingMode As afProcessingModes)
   Select Case afeProcessingMode
      Case afProcessingModeGlobalsOnly
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
Public Sub afEndProcessingMode(ByVal afeProcessingMode As afProcessingModes)
   Select Case afeProcessingMode
      Case afProcessingModeGlobalsOnly
         'Do nothing
'>>>>>>> Your cases here
         
'<<<<<<<
      Case Else
         'Do nothing
   End Select
End Sub

