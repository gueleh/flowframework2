Attribute VB_Name = "DEV_f_pM_Utilities"
' -------------------------------------------------------------------------------------------
' CORE-DEV, do not change, only required for development
'============================================================================================
' CORE-DEV - do not change, optionally remove when deploying app
'============================================================================================
'   NAME:     DEV_f_pM_Utilities
'============================================================================================
'   Purpose:  utilities for development that do require other dev resources
'   Access:   Private
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
'   0.12.1    31.01.2023    gueleh    Added Option Private Module to the module
'   0.11.0    05.08.2022    gueleh    Initially created
'--------------------------------------------------------------------------------------------
'   BACKLOG
'   ''''''''''''''''''''
'   none
'============================================================================================
Option Explicit
Option Private Module

Private Const s_m_COMPONENT_NAME As String = "DEV_f_pM_Utilities"

Public Sub DEV_Reset_DEV_f_wks_TestCanvas()
   DEV_f_wks_TestCanvas.Cells.Delete
   DEV_f_wks_TestCanvas.Rows.AutoFit
End Sub

Public Sub DEV_SetName_ScopeWorksheet()
   On Error Resume Next
   Dim oWks As Worksheet
   Set oWks = ActiveSheet
   oWks.Names.Add ActiveCell.Value2, ActiveCell.Offset(, -1)
End Sub

Public Sub DEV_SetName_ScopeWorkbook()
   On Error Resume Next
   ThisWorkbook.Names.Add ActiveCell.Value2, ActiveCell.Offset(, -1)
End Sub

