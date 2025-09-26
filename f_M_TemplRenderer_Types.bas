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

Public Type CellSpec2
    templateText As String      ' Zellliteral aus Template (mit statischen Teilen)
    PlaceholderList As String   ' "Key1|Key2|..." (alle {{...}} in dieser Zelle)
    StyleToken As String        ' z. B. Money / Body / TH
    relRow As Long              ' Zeile relativ zur Lane (1-basiert)
    relCol As Long              ' Spalte relativ zur Lane (1-basiert)
End Type

Public Type LaneSpec
    laneType As String          ' "FIX" | "REP" | "REL"
    key As String               ' LaneKey
    TopRel As Long              ' Startzeile relativ zum Block (1-basiert)
    LeftRel As Long             ' Startspalte relativ zum Block (1-basiert)
    RowsCount As Long           ' Höhe der Lane (Zeilen)
    ColsCount As Long           ' Breite der Lane (Spalten)
    Cells() As CellSpec2
    PadAfterRows As Long        ' Anzahl Leerzeilen NACH dieser Lane (REP/REL)
End Type

Public Type BlockSpec2
    blockKey As String
    ws As Worksheet
    Top As Long                 ' Block-Top (absolute Zeile im Template-Sheet)
    Left As Long                ' Block-Left (absolute Spalte im Template-Sheet)
    RowsCount As Long
    ColsCount As Long
    lanes() As LaneSpec         ' 0..n FIX + 0..n REP + 0..n REL
End Type

Public Type NamedRangeRef
    name As String
    ws As Worksheet
    rng As Range
End Type

