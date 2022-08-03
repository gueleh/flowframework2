Attribute VB_Name = "devfpMUtilities"
' -------------------------------------------------------------------------------------------
' CORE-DEV, do not change, remove when deploying
'============================================================================================
'   NAME:     devfpMUtilities
'============================================================================================
'   Purpose:  utilities for development work
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
'   0.8.0    03.08.2022    gueleh    Initially created
'--------------------------------------------------------------------------------------------
'   BACKLOG
'   ''''''''''''''''''''
'   none
'============================================================================================
Option Explicit
Option Private Module

Private Const msCOMPONENT_NAME As String = "devfpMUtilities"

Public Sub DEV_f_SetTechnicalNamesVisibleToFalse()
   Dim oC As New fCSettings
   oC.SetNamesVisibleTo False
End Sub

Public Sub DEV_f_SetTechnicalNamesVisibleToTrue()
   Dim oC As New fCSettings
   oC.SetNamesVisibleTo True
End Sub

