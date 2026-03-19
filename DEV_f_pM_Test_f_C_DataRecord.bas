Attribute VB_Name = "DEV_f_pM_Test_f_C_DataRecord"
' CORE-DEV - do not change, optionally remove when deploying app
'============================================================================================
'   NAME:     DEV_f_pM_Test_f_C_DataRecord
'============================================================================================
'   Purpose:  Unit tests for f_C_DataRecord class
'   Access:   Private
'   Type:     Module
'   Author:   Guenther Lehner
'   Contact:  guenther.lehner@protonmail.com
'   GitHubID: gueleh
'   Required: DEV_a_wks_TestCanvas, f_C_DataRecord, f_I_DataRecord,
'             DEV_f_pM_Testing, DEV_f_pM_TestRegistry
'   Usage:    Tests are registered via DEV_f_p_RegisterTests_f_C_DataRecord and run by the
'             test runner.
'--------------------------------------------------------------------------------------------
'   VERSION HISTORY
'   Version    Date    Developer    Changes
'   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   0.2.0    19.03.2026    Claude Code    Migrated to assertion framework
'   0.1.0    gueleh    Initially created
'--------------------------------------------------------------------------------------------
'   BACKLOG
'   ''''''''''''''''''''
'   none
'============================================================================================
Option Explicit
Option Private Module

Private Const s_m_COMPONENT_NAME As String = "DEV_f_pM_Test_f_C_DataRecord"

' Purpose: registers all test subs of this module with the test registry
' 0.2.0    19.03.2026    Claude Code    Initially created
Public Sub DEV_f_p_RegisterTests_f_C_DataRecord()
   DEV_f_p_AddTestSub "DEV_f_pM_Test_f_C_DataRecord.mTest_SetAndGetFieldValues"
   DEV_f_p_AddTestSub "DEV_f_pM_Test_f_C_DataRecord.mTest_GetNonExistentField"
End Sub

' Purpose: tests that field values can be set and retrieved correctly
' 0.2.0    19.03.2026    Claude Code    Migrated from Debug.Assert to assertion framework
Public Sub mTest_SetAndGetFieldValues()
   Dim oC_Test As DEV_f_C_UnitTest
   Set oC_Test = oC_DEV_f_p_CreateTest( _
      "DataRecord stores and retrieves field values correctly", _
      s_m_COMPONENT_NAME, "mTest_SetAndGetFieldValues")

   Dim oCWks As New f_C_Wks
   Dim oCRecord As f_I_DataRecord
   Dim lRow As Long
   Dim lColumn As Long
   Dim vValue As Variant

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

   Set oCRecord = New f_C_DataRecord
   For lColumn = 1 To 10
      oCRecord.bSetFieldValue oCWks.oWks_prop_r.Cells(1, lColumn), oCWks.oWks_prop_r.Cells(2, lColumn)
   Next lColumn

   Dim bResult As Boolean
   bResult = oCRecord.bGetFieldValue("$B$1", vValue)

   oC_Test.oC_prop_r_Assert.AssertTrue bResult, _
      "bGetFieldValue should return True for existing field '$B$1'"
   oC_Test.oC_prop_r_Assert.AssertEqual "$B$2", CStr(vValue), _
      "Value for field '$B$1' should be '$B$2'"

   DEV_f_p_CompleteTest oC_Test
End Sub

' Purpose: tests that getting a non-existent field returns False
' 0.2.0    19.03.2026    Claude Code    Migrated from Debug.Assert to assertion framework
Public Sub mTest_GetNonExistentField()
   Dim oC_Test As DEV_f_C_UnitTest
   Set oC_Test = oC_DEV_f_p_CreateTest( _
      "DataRecord returns False for non-existent field names", _
      s_m_COMPONENT_NAME, "mTest_GetNonExistentField")

   Dim oCWks As New f_C_Wks
   Dim oCRecord As f_I_DataRecord
   Dim lRow As Long
   Dim lColumn As Long
   Dim vValue As Variant

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

   Set oCRecord = New f_C_DataRecord
   For lColumn = 1 To 10
      oCRecord.bSetFieldValue oCWks.oWks_prop_r.Cells(1, lColumn), oCWks.oWks_prop_r.Cells(2, lColumn)
   Next lColumn

   oC_Test.oC_prop_r_Assert.AssertFalse oCRecord.bGetFieldValue("$B$20", vValue), _
      "bGetFieldValue should return False for non-existent field '$B$20'"

   DEV_f_p_CompleteTest oC_Test
End Sub


