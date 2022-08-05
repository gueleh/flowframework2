Attribute VB_Name = "devfpMUtilities"
' -------------------------------------------------------------------------------------------
' CORE-DEV, do not change, only required for development
'============================================================================================
'   NAME:     devfpMUtilities
'============================================================================================
'   Purpose:  utilities for development that do require other dev resources
'   Access:   Public
'   Type:     Module
'   Author:   Günther Lehner
'   Contact:  guleh@pm.me
'   GitHubID: gueleh
'   Required:
'   Usage:
'--------------------------------------------------------------------------------------------
'   VERSION HISTORY
'   Version    Date    Developer    Changes
'   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   0.11.0    05.08.2022    gueleh    Initially created
'--------------------------------------------------------------------------------------------
'   BACKLOG
'   ''''''''''''''''''''
'   none
'============================================================================================
Option Explicit

Private Const s_m_COMPONENT_NAME As String = "devfpMUtilities"

Public Sub DEV_Reset_devfwksTestCanvas()
   devfwksTestCanvas.Cells.Delete
   devfwksTestCanvas.Rows.AutoFit
End Sub

