Attribute VB_Name = "f_M_TemplRendererLite_Types"
' CORE, do not change
'============================================================================================
'   NAME:     f_M_TemplRendererLite_Types
'============================================================================================
'   Purpose:  public types for the Template Renderer Lite functionality
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
'--------------------------------------------------------------------------------------------
'   BACKLOG
'   ''''''''''''''''''''
'   none
'============================================================================================

Option Explicit

Private Const s_m_COMPONENT_NAME As String = "f_M_TemplRendererLite_Types"

Public Type u_f_TRlite_CellSpec
    sTemplateText As String   ' kompletter Template-Literal (inkl. statischer Teile)
    sPlaceholderList As String ' "Invoice.Number|Customer.Name|Items[i].Qty" ...
    sStyleToken As String     ' z. B. Money / Body / TH
    lRelColInBlock As Long           ' relative Spalte im Block
End Type

Public Type u_f_TRlite_RowSpec
    bIsRepeater As Boolean
    uaCellspecs() As u_f_TRlite_CellSpec
End Type

Public Type u_f_TRlite_BlockSpec
    lTop As Long
    lLeft As Long
    uaRowSpecs() As u_f_TRlite_RowSpec
    lWidth As Long
End Type

