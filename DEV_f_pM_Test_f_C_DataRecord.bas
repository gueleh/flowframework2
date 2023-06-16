Attribute VB_Name = "DEV_f_pM_Test_f_C_DataRecord"
' CORE-DEV - do not change, optionally remove when deploying app
'============================================================================================
'   NAME:     DEV_f_pM_Test_f_C_DataRecord
'============================================================================================
'   Purpose:  directly accessible dev helpers
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
Option Private Module

Private Const s_m_COMPONENT_NAME As String = "DEV_f_pM_Test_f_C_DataRecord"

Private Sub mTest_f_C_DataRecord()
   Dim oCWks As New f_C_Wks
   Dim oCRecord As f_I_DataRecord
   Dim lRow As Long
   Dim lColumn As Long
   Dim vValue As Variant
   
   'Seed test data
   With oCWks
      .Construct DEV_a_wks_TestCanvas
      .DeleteAllContents
      With .oWks_prop_r
         For lRow = 1 To 10
            For lColumn = 1 To 10
               .Cells(lRow, lColumn) = .Cells(lRow, lColumn).Address
            Next lColumn
         Next lRow
      End With
   End With
   
   'Create data record for row 2 with row 1 as header
   Set oCRecord = New f_C_DataRecord
   For lColumn = 1 To 10
      If Not oCRecord.bSetFieldValue(oCWks.oWks_prop_r.Cells(1, lColumn), oCWks.oWks_prop_r.Cells(2, lColumn)) Then
         'Error
      End If
   Next lColumn
   
   Debug.Assert oCRecord.bGetFieldValue("$B$1", vValue) = True
   Debug.Assert vValue = "$B$2"
   Debug.Assert oCRecord.bGetFieldValue("$B$20", vValue) = False
End Sub

