Attribute VB_Name = "f_pM_Globals"
' -------------------------------------------------------------------------------------------
' CORE - do not change
'============================================================================================
'   NAME:     f_pM_Globals
'============================================================================================
'   Purpose:  the globals of the framework
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
'   0.1.0    17.03.2023    gueleh    Initially created
'--------------------------------------------------------------------------------------------
'   BACKLOG
'   ''''''''''''''''''''
'   none
'============================================================================================
Option Explicit
Option Private Module

Private Const s_m_COMPONENT_NAME As String = "f_pM_Globals"

Public Const s_f_p_FRAMEWORK_VERSION_NUMBER As String = "0.3.0"
Public Const dte_f_p_FRAMEWORK_VERSION_DATE As Date = "21.03.2023"

Public Const l_f_p_EXECUTION_ERROR_NUMBER As Long = 9999
Public Const s_f_p_EXECUTION_ERROR_TEXT As String = "The execution of the called function failed."
Public Const s_f_p_ERROR As String = "<ERROR>"

