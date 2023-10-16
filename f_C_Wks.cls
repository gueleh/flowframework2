VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "f_C_Wks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' CORE, do not change
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
'     + see the sub which is testing the class in DEV_f_pM_Testing
'     + this class can be used independently from the framework
'--------------------------------------------------------------------------------------------
'   VERSION HISTORY
'   Version    Date    Developer    Changes
'   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   0.2.0    20.03.2023    gueleh    Initially created
'--------------------------------------------------------------------------------------------
'   BACKLOG
'   ''''''''''''''''''''
'TODO: [+] f_C_Wks:  add required events for oWks_m_Wks
'============================================================================================
Option Explicit

Private Const s_m_COMPONENT_NAME As String = "f_C_Wks"

Private WithEvents oWks_m_Wks As Worksheet
Attribute oWks_m_Wks.VB_VarHelpID = -1
Private oRng_m_Data As Range
Private b_m_FirstRowIsHeaderRow As Boolean
Private oDict_m_ColumnNumbersByHeaderName As Scripting.Dictionary
Private s_m_NameOfSubToRunOnChange As String
Private b_m_WksChangeEventIsActive As Boolean

Private Sub Class_Terminate()
   Set oDict_m_ColumnNumbersByHeaderName = Nothing
End Sub

Public Property Let b_prop_rw_WksChangeEventIsActive(ByVal bNewValue As Boolean)
   b_m_WksChangeEventIsActive = bNewValue
End Property

Public Property Get b_prop_rw_WksChangeEventIsActive() As Boolean
   b_prop_rw_WksChangeEventIsActive = b_m_WksChangeEventIsActive
End Property

Public Property Let s_prop_rw_NameOfSubToRunOnWksChange(ByVal sName As String)
   s_m_NameOfSubToRunOnChange = sName
End Property

Public Property Get s_prop_rw_NameOfSubToRunOnWksChange() As String
   s_prop_rw_NameOfSubToRunOnWksChange = s_m_NameOfSubToRunOnChange
End Property

Public Property Get oWks_prop_r() As Worksheet
   If b_m_WksIsNotSet() Then Exit Property
   Set oWks_prop_r = oWks_m_Wks
End Property

Public Property Get oRng_prop_r_Data() As Range
   If b_m_DataRangeIsNotSet() Then Exit Property
   Set oRng_prop_r_Data = oRng_m_Data
End Property

Public Property Get oRng_prop_r_DataWithoutHeader() As Range
   If b_m_DataRangeIsNotSet() Then Exit Property
   Dim lOffset As Long
   If b_m_FirstRowIsHeaderRow Then lOffset = 1
   Set oRng_prop_r_DataWithoutHeader = oRng_m_Data.Offset(1).Resize(oRng_m_Data.Rows.Count - lOffset)
End Property

Public Property Get l_prop_r_ColumnNumberByHeaderName(ByVal sHeaderName As String) As Long
   If oDict_m_ColumnNumbersByHeaderName Is Nothing Then Exit Property
   If oDict_m_ColumnNumbersByHeaderName.Exists(sHeaderName) Then
      l_prop_r_ColumnNumberByHeaderName = oDict_m_ColumnNumbersByHeaderName(sHeaderName)
   End If
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

' Purpose: sets the data range based on anchor cell(s)
' Usage:
'  + without top right anchor cell the range is determined based on the end of current region
'  + if first row of the data range is the header row, then please set the boolean to true
'  + if first row of the data range is the header row and you immediately want to create
'     the dictionary with column numbers by header names, then please set the boolean to true
' 0.3.0    21.03.2023    gueleh    Initially created
Public Sub SetDataRangeByAnchors(ByRef oRngAnchorCellTopLeft As Range, _
   Optional ByRef oRngAnchorCellTopRight As Range, _
   Optional ByVal bFirstRowIsHeaderRow As Boolean = False, _
   Optional ByVal bCreateHeaderDictionary As Boolean = False)

   If b_m_WksIsNotSet() Then Exit Sub
   If oRngAnchorCellTopRight Is Nothing Then
      If b_m_ParentIsDifferent(oRngAnchorCellTopLeft) Then Exit Sub
      Set oRng_m_Data = oRng_m_SanitizedCurrentRegion(oRngAnchorCellTopLeft)
   Else
      If b_m_ParentIsDifferent(oRngAnchorCellTopLeft, oRngAnchorCellTopRight) Then Exit Sub
      Set oRng_m_Data = oRng_m_SanitizedCurrentRegion(oRngAnchorCellTopLeft, oRngAnchorCellTopRight)
   End If

   If bFirstRowIsHeaderRow Then
      b_m_FirstRowIsHeaderRow = True
      If bCreateHeaderDictionary Then CreateHeaderDictionary
   End If

End Sub

Public Sub CreateHeaderDictionary(Optional ByVal lRowWithHeaders As Long = 0)
   If b_m_WksIsNotSet() Then Exit Sub
   If b_m_DataRangeIsNotSet() Then Exit Sub
   If b_m_HeaderCannotBeSet(lRowWithHeaders) Then Exit Sub
   
   Dim oRngHeaderRow As Range
   Dim oRngHeaderCell As Range
   Dim sHeaderName As String
   Dim lHeaderRow As Long
   
   If lRowWithHeaders > 0 Then
      lHeaderRow = lRowWithHeaders
   Else
      lHeaderRow = oRng_m_Data.Row
   End If
   
   Set oRngHeaderRow = oWks_m_Wks.Range( _
      oWks_m_Wks.Cells(lHeaderRow, oRng_m_Data.Column), _
      oWks_m_Wks.Cells(lHeaderRow, (oRng_m_Data.Column + oRng_m_Data.Columns.Count - 1)))
      
   Set oDict_m_ColumnNumbersByHeaderName = New Scripting.Dictionary
   For Each oRngHeaderCell In oRngHeaderRow
      sHeaderName = oRngHeaderCell.Value2
      If sHeaderName = "" Then GoTo Catch
      If Not oDict_m_ColumnNumbersByHeaderName.Exists(sHeaderName) Then
         oDict_m_ColumnNumbersByHeaderName.Add sHeaderName, oRngHeaderCell.Column
      Else
         GoTo Catch
      End If
   Next oRngHeaderCell
      
   Exit Sub
Catch:
   MsgBox "Header could not be set due to invalid data in header row, i.e. names not unique or cells empty", vbCritical
   Set oDict_m_ColumnNumbersByHeaderName = Nothing
End Sub

Private Function b_m_HeaderCannotBeSet(ByVal lRow As Long) As Boolean
   Dim bCannotBeSet As Boolean
   If lRow = 0 And Not b_m_FirstRowIsHeaderRow Then
      bCannotBeSet = True
   ElseIf lRow > 0 _
   And lRow >= oRng_m_Data.Row _
   And lRow <= (oRng_m_Data.Row + oRng_m_Data.Rows.Count - 1) _
   Then
      bCannotBeSet = True
   ElseIf lRow >= (oWks_m_Wks.UsedRange.Row + oWks_m_Wks.UsedRange.Rows.Count - 1) _
   Then
      bCannotBeSet = True
   End If
   
   If bCannotBeSet Then
      MsgBox "Header cannot be set - either first row of data range has to be specified as header row when setting the data range or a valid row number outside of the data range has to be provided.", vbCritical
      b_m_HeaderCannotBeSet = True
   End If
End Function

' Purpose: sanitizes the used range to cover only the really used range, fails if it contains merged cells
' 0.3.0    21.03.2023    gueleh    Initially created
Public Sub SanitizeUsedRange(Optional ByVal bFailSilentlyInCaseOfMergedCells = False)
   If b_m_WksIsNotSet() Then Exit Sub
   If b_m_UsedRangeHasMergedCells(bFailSilentlyInCaseOfMergedCells) Then Exit Sub

   Dim lLastRow As Long
   Dim lLastColumn As Long

   With oWks_m_Wks
'        Set dummyRng = .UsedRange
      On Error Resume Next
      lLastRow = .Cells.Find("*", After:=.Cells(1), LookIn:=xlFormulas, Lookat:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
      lLastColumn = .Cells.Find("*", After:=.Cells(1), LookIn:=xlFormulas, Lookat:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
      On Error GoTo 0

      If lLastRow * lLastColumn = 0 Then
          DeleteAllContents
      Else
          .Range(.Cells(lLastRow + 1, 1), _
                 .Cells(.Rows.Count, 1)).EntireRow.Delete
          .Range(.Cells(1, lLastColumn + 1), _
                 .Cells(1, .Columns.Count)).EntireColumn.Delete
      End If
    End With
End Sub

Private Function b_m_UsedRangeHasMergedCells(ByVal bFailSilently As Boolean) As Boolean
   If oWks_m_Wks.UsedRange.MergeCells Then
      MsgBox "The worksheet " & oWks_m_Wks.Name & " has merged cells. Please unmerge all cells and try again.", vbCritical
      b_m_UsedRangeHasMergedCells = True
   End If
End Function

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

Private Sub oWks_m_Wks_Change(ByVal Target As Range)
   If Not b_m_WksChangeEventIsActive Then Exit Sub
   If Len(s_m_NameOfSubToRunOnChange) = 0 Then Exit Sub
   Application.Run "'" & ThisWorkbook.Name & "'!" & s_m_NameOfSubToRunOnChange, Target
End Sub
