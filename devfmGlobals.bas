Attribute VB_Name = "devfmGlobals"
' -------------------------------------------------------------------------------------------
' CORE-DEV, do not change
'============================================================================================
'   NAME:     devfmGlobals
'============================================================================================
'   Purpose:  the core globals when developing, can be removed along with all other modules when deploying
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

Private Const msCOMPONENT_NAME As String = "devfmGlobals"

' Purpose: initializes the globals which are required for development work
' 0.1.0    20220709    gueleh    Initially created
Public Sub devfInitGlobals()
   Set fgcolUnitTests = New Collection
End Sub


