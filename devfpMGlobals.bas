Attribute VB_Name = "devfpMGlobals"
' -------------------------------------------------------------------------------------------
' CORE-DEV, do not change
'============================================================================================
'   NAME:     devfpMGlobals
'============================================================================================
'   Purpose:  the core globals when developing, can be removed along with all other modules when deploying
'   Access:   Private
'   Type:     Module
'   Author:   G?nther Lehner
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

Private Const s_m_COMPONENT_NAME As String = "devfpMGlobals"

' Purpose: initializes the globals which are required for development work
' 0.1.0    20220709    gueleh    Initially created
Public Sub DEV_f_g_InitGlobals()
   Set oCol_f_g_UnitTests = New Collection
End Sub


