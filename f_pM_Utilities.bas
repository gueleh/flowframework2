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

' Purpose: Return string for Application.Run for procedures in this workbook - if required, MyProcedureName with Module Selector, e.g. "MyModule.MyProcedureName"
' 0.1.0    20220709    gueleh    Initially created
Public Function s_f_p_MyProcedureName _
( _
   ByVal s_arg_ProcedureName As String, _
   Optional ByRef oWkb_arg_Workbook As Workbook _
) As String
  s_f_p_MyProcedureName = "'" & oWkb_f_p_DefaultToThisWorkbook(oWkb_arg_Workbook).name & "'!" & s_arg_ProcedureName
End Function

' Purpose: Return provided Workbook object or as a default ThisWorkbook - a helper to facilitate solutions with more than one workbook involved
Public Function oWkb_f_p_DefaultToThisWorkbook(Optional ByRef oWkb As Workbook)
   If oWkb Is Nothing Then
      Set oWkb_f_p_DefaultToThisWorkbook = ThisWorkbook
   Else
      Set oWkb_f_p_DefaultToThisWorkbook = oWkb
   End If
End Function

' Purpose: adds ' to all array items starting with 0
' 0.11.0    05.08.2022    gueleh    Initially created
Public Sub f_p_SanitizeLeadingZeroItems _
( _
   ByRef va_arg_Data() As Variant _
)
   
   Dim lRow As Long
   Dim lColumn As Long
   On Error Resume Next 'make sure to process all that can be processed
   For lRow = LBound(va_arg_Data, 1) To UBound(va_arg_Data, 1)
      For lColumn = LBound(va_arg_Data, 2) To UBound(va_arg_Data, 2)
         If Left$(va_arg_Data(lRow, lColumn), 1) = "0" Then
            va_arg_Data(lRow, lColumn) = "'" & va_arg_Data(lRow, lColumn)
         End If
      Next lColumn
   Next lRow
End Sub

' Purpose: returns value from worksheet-scope named cell (empty if error)
' 0.1.0    17.03.2023    gueleh    Initially created
Public Function v_f_p_ValueFromWorksheetName _
( _
   ByRef o_arg_Wks As Worksheet, _
   ByVal s_arg_Name As String _
) As Variant
   
   On Error Resume Next
   v_f_p_ValueFromWorksheetName = o_arg_Wks.Names(s_arg_Name).RefersToRange.Value2
   If Err.Number > 0 Then v_f_p_ValueFromWorksheetName = s_f_p_ERROR
End Function

' Purpose: returns range from worksheet-scope named range (nothing if error)
' 0.1.0    17.03.2023    gueleh     Initially created
Public Function oRng_f_p_RangeFromWorksheetName _
( _
   ByRef o_arg_Wks As Worksheet, _
   ByVal s_arg_Name As String _
) As Range
   
   On Error Resume Next
   Set oRng_f_p_RangeFromWorksheetName = o_arg_Wks.Names(s_arg_Name).RefersToRange
End Function

' Purpose: provide CodeName of Worksheet as String and optionally Workbook, get Worksheet object
Public Function oWks_f_p_WorksheetFromCodeNameString(ByVal s_arg_CodeName As String, Optional ByRef oWkb_arg_Workbook As Workbook)
   Dim oWkb As Workbook
   Dim oWks As Worksheet

   On Error Resume Next
   Set oWkb = oWkb_f_p_DefaultToThisWorkbook(oWkb_arg_Workbook)
   For Each oWks In oWkb.Worksheets
      If oWks.CodeName = s_arg_CodeName Then
         Set oWks_f_p_WorksheetFromCodeNameString = oWks
         Exit Function
      End If
   Next oWks
   
End Function

' Purpose: provide Name of Worksheet as String and optionally Workbook, get Worksheet object
Public Function oWks_f_p_WorksheetFromName(ByVal s_arg_Name As String, Optional ByRef oWkb_arg_Workbook As Workbook)
   Dim oWkb As Workbook
   Dim oWks As Worksheet

   On Error Resume Next
   Set oWkb = oWkb_f_p_DefaultToThisWorkbook(oWkb_arg_Workbook)
   For Each oWks In oWkb.Worksheets
      If oWks.name = s_arg_Name Then
         Set oWks_f_p_WorksheetFromName = oWks
         Exit Function
      End If
   Next oWks
   
End Function


' Purpose: returns value from worksheet-scope named cell (empty if error)
' 0.1.0    17.03.2023    gueleh    Initially created
Public Function v_f_p_ValueFromWorkbookName _
( _
   ByVal s_arg_Name As String, _
   Optional ByRef oWkb_arg_Workbook As Workbook _
) As Variant
   
   On Error Resume Next
   v_f_p_ValueFromWorkbookName = oWkb_f_p_DefaultToThisWorkbook(oWkb_arg_Workbook).Names(s_arg_Name).RefersToRange.Value2
   If Err.Number > 0 Then v_f_p_ValueFromWorkbookName = s_f_p_ERROR
End Function

' Purpose: returns range from worksheet-scope named range (nothing if error)
' 0.1.0    17.03.2023    gueleh     Initially created
Public Function oRng_f_p_RangeFromWorkbookName _
( _
   ByVal s_arg_Name As String, _
   Optional ByRef oWkb_arg_Workbook As Workbook _
) As Range
   
   On Error Resume Next
   Set oRng_f_p_RangeFromWorkbookName = oWkb_f_p_DefaultToThisWorkbook(oWkb_arg_Workbook).Names(s_arg_Name).RefersToRange
End Function

' Purpose: sanitizes a numeric key so that it certainly works with dictionaries
' 1.8.0    13.11.2023    gueleh    Initially created
Public Function s_f_p_SanitizedKey _
( _
   ByVal v_arg_Key As Variant _
) As String
   
   s_f_p_SanitizedKey = s_f_p_SPLIT_SEED_SEPARATOR & CStr(v_arg_Key)
End Function

' Purpose: restores the Long value from the string of a sanitized key
' 1.8.0    13.11.2023    gueleh    Initially created
Public Function l_f_p_KeyFromSanitizedKey _
( _
   ByVal s_arg_Key As String _
) As String
   
   l_f_p_KeyFromSanitizedKey = CLng(Replace$(s_arg_Key, s_f_p_SPLIT_SEED_SEPARATOR, ""))
End Function
