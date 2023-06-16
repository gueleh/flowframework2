Attribute VB_Name = "DEV_f_pM_Globals"
' CORE-DEV - do not change, optionally remove when deploying app
'============================================================================================
'   NAME:     DEV_f_pM_Globals
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

Private Const s_m_COMPONENT_NAME As String = "DEV_f_pM_Globals"

' Purpose: initializes the globals which are required for development work
' 0.1.0    20220709    gueleh    Initially created
Public Sub DEV_f_p_InitGlobals()
   Set oCol_f_p_UnitTests = New Collection
End Sub


