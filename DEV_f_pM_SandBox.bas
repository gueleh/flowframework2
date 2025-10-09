Attribute VB_Name = "DEV_f_pM_SandBox"
' CORE-DEV - do not change, optionally remove when deploying app
'============================================================================================
'   NAME:     DEV_f_pM_SandBox
'============================================================================================
'   Purpose: Sandbox for dev experiments
'   Access:   Private
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
Option Private Module

Private Const s_m_COMPONENT_NAME As String = "DEV_f_pM_SandBox"

Private Sub mTestReadingWorkbookMarker()
   Debug.Print b_f_p_WorkbookHasFlowFrameworkMarker()
End Sub

Private Sub mSetAndCheckFlowFrameworkMarkerForSelectedWorkbook()
   Dim fd As FileDialog
   Dim filePath As Variant
   Dim wb As Workbook
   Dim added As Boolean
   Dim hasMarker As Boolean

   On Error GoTo CleanFail

   Set fd = Application.FileDialog(msoFileDialogFilePicker)
   With fd
      .Title = "Workbook auswählen"
      .AllowMultiSelect = False
      .Filters.Clear
      .Filters.Add "Excel-Arbeitsmappen", "*.xlsx;*.xlsm;*.xlsb;*.xls"
      If .Show <> -1 Then
         Debug.Print "Aktion abgebrochen."
         Exit Sub
      End If
      filePath = .SelectedItems(1)
   End With

   Set wb = Workbooks.Open(CStr(filePath), ReadOnly:=False)

   added = b_f_p_AddFlowFrameworkMarker(wb)               ' Marker setzen (nur wenn noch nicht vorhanden)
   hasMarker = b_f_p_WorkbookHasFlowFrameworkMarker(wb)   ' Marker prüfen

   Debug.Print "Datei: "; wb.FullName
   Debug.Print "Marker neu gesetzt: "; added
   Debug.Print "Marker vorhanden: "; hasMarker

   If added Then
      wb.Save
      Debug.Print "Workbook gespeichert (Marker persistiert)."
   End If

   Exit Sub

CleanFail:
   Debug.Print "Fehler (" & Err.Number & "): " & Err.Description
End Sub


' Purpose: tests zero sanitation manually, if the code executes without stopping the tests were successful
' 0.11.0    05.08.2022    gueleh    Initially created
Private Sub mManualTest_RangeArrayProcessorZeroSanitation()
   Dim oC As New f_C_RangeArrayProcessor
   Dim va() As Variant
   DEV_Reset_DEV_f_wks_TestCanvas
   With DEV_f_wks_TestCanvas
      .Range("A1").Value = "ID"
      .Range("B1").Value = "ID2"
      .Range("C1").Value = "Value"
      .Range("A2").Value = "'01"
      .Range("B2").Value = "AgA"
      .Range("C2").Value = "What is AgA?"
      .Columns.AutoFit
      va = .Range("A1").CurrentRegion.Formula
      .Range("A1").CurrentRegion.Formula = va
      Debug.Assert .Range("A2").Value = "1"
      oC.SanitizeLeadingZeroItems va
      .Range("A1").CurrentRegion.Formula = va
      Debug.Assert .Range("A2").Value = "01"
   End With
   DEV_Reset_DEV_f_wks_TestCanvas
End Sub
