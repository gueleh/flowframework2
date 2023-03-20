VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "f_C_Wks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' -------------------------------------------------------------------------------------------
' FRAMEWORK - do not change
'============================================================================================
'   NAME:     f_C_Wks
'============================================================================================
'   Purpose:  enhanced worksheet class
'   Access:   Public
'   Type:     Class Module
'   Author:   Günther Lehner
'   Contact:  guleh@pm.me
'   GitHubID: gueleh
'   Required:
'   Usage:
'--------------------------------------------------------------------------------------------
'   VERSION HISTORY
'   Version    Date    Developer    Changes
'   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   0.2.0    20.03.2023    gueleh    Initially created
'--------------------------------------------------------------------------------------------
'   BACKLOG
'   ''''''''''''''''''''
'   none
'============================================================================================
Option Explicit

Private Const s_m_COMPONENT_NAME As String = "f_C_Wks"

Private oWks_m_Wks As Worksheet
Attribute oWks_m_Wks.VB_VarHelpID = -1
Private oRng_m_Data As Range

Public Property Get oWks_prop_r() As Worksheet
   If b_m_WksIsNotSet() Then Exit Property
   Set oWks_prop_r = oWks_m_Wks
End Property

Public Property Get oRng_prop_rw_Data() As Range
   If b_m_DataRangeIsNotSet() Then Exit Property
   Set oRng_prop_rw_Data = oRng_m_Data
End Property

' Purpose: gets current range under consideration of anchor cell(s), i.e. ignoring potential
'        cells above, left of and right of the anchor cells(s)
' Usage:
'  + top left anchor should be the top left cell of the range you want to get - cells above and left of it are ignored
'  + optionally provide top right anchor, which is only relevant for the number of colums to be included
'  + if optional anchor not provided, number of columns is determined automatically
'  + refer to the test sub in DEV_f_pM_Testing for an example usage
' 0.2.0    20.03.2023    gueleh    Initially created
Public Property Get oRng_prop_r_CurrentRegionEnhanced( _
   ByRef oRngAnchorCellTopLeft As Range, _
   Optional ByRef oRngAnchorCellTopRight As Range) As Range
   
   If oRngAnchorCellTopRight Is Nothing Then
      Set oRng_prop_r_CurrentRegionEnhanced = oRng_m_SanitizedCurrentRegion(oRngAnchorCellTopLeft)
   Else
      Set oRng_prop_r_CurrentRegionEnhanced = oRng_m_SanitizedCurrentRegion(oRngAnchorCellTopLeft, oRngAnchorCellTopRight)
   End If
   
End Property

Public Sub Construct(ByRef oWks As Worksheet)
   Set oWks_m_Wks = oWks
End Sub

Public Sub DeleteAllContents()
   If b_m_WksIsNotSet() Then Exit Sub
   oWks_m_Wks.UsedRange.Rows.Delete
End Sub

Public Sub SetDataRangeByAnchors(ByRef oRngAnchorCellTopLeft As Range, _
   Optional ByRef oRngAnchorCellTopRight As Range)

   If b_m_WksIsNotSet() Then Exit Sub
   If oRngAnchorCellTopRight Is Nothing Then
      If b_m_ParentIsDifferent(oRngAnchorCellTopLeft) Then Exit Sub
      Set oRng_m_Data = oRng_m_SanitizedCurrentRegion(oRngAnchorCellTopLeft)
   Else
      If b_m_ParentIsDifferent(oRngAnchorCellTopLeft, oRngAnchorCellTopRight) Then Exit Sub
      Set oRng_m_Data = oRng_m_SanitizedCurrentRegion(oRngAnchorCellTopLeft, oRngAnchorCellTopRight)
   End If

End Sub

Public Sub SanitizeUsedRange()
   'TODO: [+] SanitizeUsedRange:  write the code for this
End Sub

Private Function b_m_WksIsNotSet() As Boolean
   If oWks_m_Wks Is Nothing Then
      MsgBox "The worksheet is not set for this instance of f_C_Wks. Please use the method Construct to set the worksheet.", vbCritical
      b_m_WksIsNotSet = True
   End If
End Function

Private Function b_m_DataRangeIsNotSet() As Boolean
   If oRng_m_Data Is Nothing Then
      MsgBox "The data range is not set for this instance of f_C_Wks. Please use the method SetDataRangeByAnchors or the property to set the data range.", vbCritical
      b_m_DataRangeIsNotSet = True
   End If
End Function

Private Function b_m_ParentIsDifferent(ParamArray aArgs() As Variant) As Boolean
   Dim oRng As Range
   Dim oWks As Worksheet
   Dim lIndex As Long
   For lIndex = LBound(aArgs) To UBound(aArgs)
      On Error Resume Next
      Set oRng = aArgs(lIndex)
      If Err.Number > 0 Then GoTo Catch
      Set oWks = oRng.Parent
      If oWks.CodeName <> oWks_m_Wks.CodeName Then GoTo Catch
   Next lIndex
   Exit Function
Catch:
   MsgBox "Parent worksheet of provided range is not the worksheet of the class instance.", vbCritical
   b_m_ParentIsDifferent = True
End Function

Private Function oRng_m_SanitizedCurrentRegion(ByRef oRngAnchorCellTopLeft As Range, _
   Optional ByRef oRngAnchorCellTopRight As Range) As Range

   Dim oRngCurrent As Range
   Dim lRowOffset As Long
   Dim lRowSize As Long
   Dim lColumnOffset As Long
   Dim lColumnSize As Long
   
   Set oRngCurrent = oRngAnchorCellTopLeft.CurrentRegion
   lRowOffset = oRngAnchorCellTopLeft.Row - oRngCurrent.Row
   lColumnOffset = oRngAnchorCellTopLeft.Column - oRngCurrent.Column
   lRowSize = oRngCurrent.Rows.Count - lRowOffset
   If oRngAnchorCellTopRight Is Nothing Then
      lColumnSize = oRngCurrent.Columns.Count - lColumnOffset
   Else
      lColumnSize = oRngAnchorCellTopRight.Column - lColumnOffset
   End If
   
   Set oRng_m_SanitizedCurrentRegion = oRngCurrent.Offset(lRowOffset, lColumnOffset).Resize(lRowSize, lColumnSize)
      

End Function

