Attribute VB_Name = "f_pM_Utilities"
' CORE, do not change
'============================================================================================
'   NAME:     f_pM_Utilities
'============================================================================================
'   Purpose:  utilities being part of the core of the template
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
'   0.1.0    20220709    gueleh    Initially created
'--------------------------------------------------------------------------------------------
'   BACKLOG
'   ''''''''''''''''''''
'   none
'============================================================================================
Option Explicit
Option Private Module

Private Const s_m_COMPONENT_NAME As String = "f_pM_Utilities"

' Purpose: Return string for Application.Run for procedures in this workbook
' 0.1.0    20220709    gueleh    Initially created
Public Function s_f_p_MyProcedureName(ByVal sProcedureName As String) As String
  s_f_p_MyProcedureName = "'" & ThisWorkbook.Name & "'!" & sProcedureName
End Function

' Purpose: adds ' to all array items starting with 0
' 0.11.0    05.08.2022    gueleh    Initially created
Public Sub f_p_SanitizeLeadingZeroItems(ByRef vaData() As Variant)
   Dim lRow As Long
   Dim lColumn As Long
   On Error Resume Next 'make sure to process all that can be processed
   For lRow = LBound(vaData, 1) To UBound(vaData, 1)
      For lColumn = LBound(vaData, 2) To UBound(vaData, 2)
         If Left$(vaData(lRow, lColumn), 1) = "0" Then vaData(lRow, lColumn) = "'" & vaData(lRow, lColumn)
      Next lColumn
   Next lRow
End Sub

' Purpose: returns value from worksheet-scope named cell (empty if error)
' 0.1.0    17.03.2023    gueleh    Initially created
Public Function v_f_p_ValueFromWorksheetName(ByRef oWks As Worksheet, ByVal sName As String) As Variant
   On Error Resume Next
   v_f_p_ValueFromWorksheetName = oWks.Names(sName).RefersToRange.Value2
   If Err.Number > 0 Then v_f_p_ValueFromWorksheetName = s_f_p_ERROR
End Function

' Purpose: returns range from worksheet-scope named range (nothing if error)
' 0.1.0    17.03.2023    gueleh     Initially created
Public Function oRng_f_p_RangeFromWorksheetName(ByRef oWks As Worksheet, ByVal sName As String) As Variant
   On Error Resume Next
   oRng_f_p_RangeFromWorksheetName = oWks.Names(sName).RefersToRange
End Function

' Purpose: returns value from worksheet-scope named cell (empty if error)
' 0.1.0    17.03.2023    gueleh    Initially created
Public Function v_f_p_ValueFromWorkbookName(ByVal sName As String) As Variant
   On Error Resume Next
   v_f_p_ValueFromWorkbookName = ThisWorkbook.Names(sName).RefersToRange.Value2
   If Err.Number > 0 Then v_f_p_ValueFromWorkbookName = s_f_p_ERROR
End Function

' Purpose: returns range from worksheet-scope named range (nothing if error)
' 0.1.0    17.03.2023    gueleh     Initially created
Public Function oRng_f_p_RangeFromWorkbookName(ByVal sName As String) As Variant
   On Error Resume Next
   oRng_f_p_RangeFromWorkbookName = ThisWorkbook.Names(sName).RefersToRange
End Function

