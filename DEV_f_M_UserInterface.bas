Attribute VB_Name = "DEV_f_M_UserInterface"
' CORE-DEV - do not change, optionally remove when deploying app
'============================================================================================
'   NAME:     DEV_f_UserInterface
'============================================================================================
'   Purpose:  directly accessible dev helpers
'   Access:   Public
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
' 0.10.0    04.08.2022    gueleh    Changed property names to meet new convention
' 0.7.0    02.08.2022    gueleh    Refactored name to match convention
'   0.2.0    20220711    gueleh    Initially created
'--------------------------------------------------------------------------------------------
'   BACKLOG
'   ''''''''''''''''''''
'   none
'============================================================================================
Option Explicit

Private Const s_m_COMPONENT_NAME As String = "DEV_f_UserInterface"

Public Sub DEV_f_g_ExportVersionControlData()
   DEV_f_p_ExportDataForVersionControl
End Sub

' Purpose: Shortcut Ctrl + Shift N, adding a worksheet scope name based on active sheet,
'  value of active cell as name, applied to cell left to active cell
' 1.9.0    22.11.2023    gueleh    Initially created
Public Sub DEV_f_g_SetName_ScopeWorksheet()
Attribute DEV_f_g_SetName_ScopeWorksheet.VB_Description = "For settings sheets: take value of active cell and add a named cell left to it with worksheet scope and the value of active cell as the name."
Attribute DEV_f_g_SetName_ScopeWorksheet.VB_ProcData.VB_Invoke_Func = "N\n14"
   DEV_f_pM_Utilities.DEV_SetName_ScopeWorksheet
End Sub

' Purpose: Shortcut Ctrl + Shift M, adding a workbook scope name,
'  value of active cell as name, applied to cell left to active cell
' 1.9.0    22.11.2023    gueleh    Initially created
Public Sub DEV_f_g_SetName_ScopeWorkbook()
Attribute DEV_f_g_SetName_ScopeWorkbook.VB_Description = "For settings sheets: take value of active cell and add a named cell left to it with workbook scope and the value of active cell as the name."
Attribute DEV_f_g_SetName_ScopeWorkbook.VB_ProcData.VB_Invoke_Func = "M\n14"
   DEV_f_pM_Utilities.DEV_SetName_ScopeWorkbook
End Sub

