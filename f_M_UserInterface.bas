Attribute VB_Name = "f_M_UserInterface"
' -------------------------------------------------------------------------------------------
' framework
'============================================================================================
'   NAME:     f_M_UserInterface
'============================================================================================
'   Purpose:  user interface for framework operations
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
'   1.10.0    23.11.2023    gueleh    Initially created
'--------------------------------------------------------------------------------------------
'   BACKLOG
'   ''''''''''''''''''''
'   none
'============================================================================================
Option Explicit

Private Const s_m_COMPONENT_NAME As String = "f_M_UserInterface"

Private Sub click_ToggleDevelopmentMode()
   f_p_ToggleDevelopmentMode
End Sub

Public Sub click_ToggleMaintenanceMode()
   f_p_ToggleMaintenanceMode
End Sub

