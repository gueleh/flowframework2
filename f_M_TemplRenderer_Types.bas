Attribute VB_Name = "f_M_TemplRenderer_Types"
' CORE, do not change
'============================================================================================
'   NAME:     f_M_TemplRenderer_Types
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

Private Const s_m_COMPONENT_NAME As String = "f_M_TemplRenderer_Types"

Public Type u_f_CellSpecRenderer
    sTemplateText As String      ' Zellliteral aus Template (mit statischen Teilen)
    sPlaceholderList As String   ' "Key1|Key2|..." (alle {{...}} in dieser Zelle)
    sStyleToken As String        ' z. B. Money / Body / TH
    lRelRow As Long              ' Zeile relativ zur Lane (1-basiert)
    lRelCol As Long              ' Spalte relativ zur Lane (1-basiert)
End Type

Public Type u_f_LaneSpecRenderer
    sLaneType As String          ' "FIX" | "REP" | "REL"
    sKey As String               ' LaneKey
    lTopRel As Long              ' Startzeile relativ zum Block (1-basiert)
    lLeftRel As Long             ' Startspalte relativ zum Block (1-basiert)
    lRowsCount As Long           ' Höhe der Lane (Zeilen)
    lColsCount As Long           ' Breite der Lane (Spalten)
    uaCells() As u_f_CellSpecRenderer
    lPadAfterRows As Long        ' Anzahl Leerzeilen NACH dieser Lane (REP/REL)
End Type

Public Type u_f_BlockSpecRenderer
    sBlockKey As String
    oWks As Worksheet
    lTop As Long                 ' Block-Top (absolute Zeile im Template-Sheet)
    lLeft As Long                ' Block-Left (absolute Spalte im Template-Sheet)
    lRowsCount As Long
    lColsCount As Long
    uaLanes() As u_f_LaneSpecRenderer         ' 0..n FIX + 0..n REP + 0..n REL
End Type

Public Type u_f_NamedRangeRefRenderer
    sName As String
    oWks As Worksheet
    oRng As Range
End Type


