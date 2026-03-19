Attribute VB_Name = "DEV_f_pM_Test_f_C_Wks"
' CORE-DEV - do not change, optionally remove when deploying app
'============================================================================================
'   NAME:     DEV_f_pM_Test_f_C_Wks
'============================================================================================
'   Purpose:  Unit tests for f_C_Wks class
'   Access:   Private
'   Type:     Module
'   Author:   Guenther Lehner
'   Contact:  guenther.lehner@protonmail.com
'   GitHubID: gueleh
'   Required: DEV_a_wks_TestCanvas, f_C_Wks, DEV_f_pM_Testing, DEV_f_pM_TestRegistry
'   Usage:    Tests are registered via DEV_f_p_RegisterTests_f_C_Wks and run by the test runner.
'             Can also be run individually via the mTest_* subs.
'--------------------------------------------------------------------------------------------
'   VERSION HISTORY
'   Version    Date    Developer    Changes
'   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   0.3.0    19.03.2026    Claude Code    Migrated to assertion framework
'   0.2.0    20.03.2023    gueleh    Initially created
'--------------------------------------------------------------------------------------------
'   BACKLOG
'   ''''''''''''''''''''
'   none
'============================================================================================
Option Explicit
Option Private Module

Private Const s_m_COMPONENT_NAME As String = "DEV_f_pM_Test_f_C_Wks"

Private b_m_ChangeEventInvoked As Boolean

' Purpose: registers all test subs of this module with the test registry
' 0.3.0    19.03.2026    Claude Code    Initially created
Public Sub DEV_f_p_RegisterTests_f_C_Wks()
   DEV_f_p_AddTestSub "DEV_f_pM_Test_f_C_Wks.mTest_Ping_Parent"
   DEV_f_p_AddTestSub "DEV_f_pM_Test_f_C_Wks.mTest_CurrentRegionEnhanced"
   DEV_f_p_AddTestSub "DEV_f_pM_Test_f_C_Wks.mTest_SetDataRangeByAnchors"
   DEV_f_p_AddTestSub "DEV_f_pM_Test_f_C_Wks.mTest_HeaderDictionary"
   DEV_f_p_AddTestSub "DEV_f_pM_Test_f_C_Wks.mTest_SanitizeUsedRange"
   DEV_f_p_AddTestSub "DEV_f_pM_Test_f_C_Wks.mTest_SetDataRangeWithHeaderDict"
   DEV_f_p_AddTestSub "DEV_f_pM_Test_f_C_Wks.mTest_WksChangeEvent"
End Sub

' Purpose: tests that ParentWorkbook returns the correct workbook
' 0.3.0    19.03.2026    Claude Code    Migrated from Debug.Assert to assertion framework
' 0.2.0    20.03.2023    gueleh    Initially created
Public Sub mTest_Ping_Parent()
   Dim oC_Test As DEV_f_C_UnitTest
   Set oC_Test = oC_DEV_f_p_CreateTest( _
      "ParentWorkbook returns correct workbook for local and external sheets", _
      s_m_COMPONENT_NAME, "mTest_Ping_Parent")

   Dim oC As f_C_Wks
   Dim oWkb As Workbook
   Dim oWks As Worksheet

   Set oC = New f_C_Wks
   oC.Construct a_wks_Main
   oC_Test.oC_prop_r_Assert.AssertEqual ThisWorkbook.name, oC.oWkb_prop_r_ParentWorkbook.name, _
      "ParentWorkbook of a_wks_Main should be ThisWorkbook"

   Set oWkb = Workbooks.Add()
   Set oWks = oWkb.Worksheets(1)
   Set oC = New f_C_Wks
   oC.Construct oWks
   oC_Test.oC_prop_r_Assert.AssertEqual oWkb.name, oC.oWkb_prop_r_ParentWorkbook.name, _
      "ParentWorkbook of external sheet should be the external workbook"

   Set oC = Nothing
   oWkb.Close False

   DEV_f_p_CompleteTest oC_Test
End Sub

' Purpose: tests that CurrentRegionEnhanced returns the correct bounded range
' 0.3.0    19.03.2026    Claude Code    Migrated from Debug.Assert to assertion framework
Public Sub mTest_CurrentRegionEnhanced()
   Dim oC_Test As DEV_f_C_UnitTest
   Set oC_Test = oC_DEV_f_p_CreateTest( _
      "CurrentRegionEnhanced returns correct range bounded by anchor cells", _
      s_m_COMPONENT_NAME, "mTest_CurrentRegionEnhanced")

   mSeedTestData

   Dim oCWks As New f_C_Wks
   oCWks.Construct DEV_a_wks_TestCanvas

   Dim oRng As Range
   Dim oRngEndColumn As Range
   Set oRng = oCWks.oWks_prop_r.Range("B3")
   Set oRngEndColumn = oCWks.oWks_prop_r.Range("D3")

   oC_Test.oC_prop_r_Assert.AssertEqual "$B$3:$D$7", _
      oCWks.oRng_prop_r_CurrentRegionEnhanced(oRng, oRngEndColumn).Address, _
      "CurrentRegionEnhanced should return $B$3:$D$7"

   DEV_f_p_CompleteTest oC_Test
End Sub

' Purpose: tests that SetDataRangeByAnchors correctly sets the data range
' 0.3.0    19.03.2026    Claude Code    Migrated from Debug.Assert to assertion framework
Public Sub mTest_SetDataRangeByAnchors()
   Dim oC_Test As DEV_f_C_UnitTest
   Set oC_Test = oC_DEV_f_p_CreateTest( _
      "SetDataRangeByAnchors sets correct data range without header", _
      s_m_COMPONENT_NAME, "mTest_SetDataRangeByAnchors")

   mSeedTestData

   Dim oCWks As New f_C_Wks
   oCWks.Construct DEV_a_wks_TestCanvas

   Dim oRng As Range
   Dim oRngEndColumn As Range
   Set oRng = oCWks.oWks_prop_r.Range("B3")
   Set oRngEndColumn = oCWks.oWks_prop_r.Range("D3")

   oCWks.SetDataRangeByAnchors oRng, oRngEndColumn
   oC_Test.oC_prop_r_Assert.AssertEqual "$B$3:$D$7", oCWks.oRng_prop_r_Data.Address, _
      "Data range should be $B$3:$D$7"

   DEV_f_p_CompleteTest oC_Test
End Sub

' Purpose: tests that CreateHeaderDictionary builds correct column mappings
' 0.3.0    19.03.2026    Claude Code    Migrated from Debug.Assert to assertion framework
Public Sub mTest_HeaderDictionary()
   Dim oC_Test As DEV_f_C_UnitTest
   Set oC_Test = oC_DEV_f_p_CreateTest( _
      "CreateHeaderDictionary maps header names to correct column numbers", _
      s_m_COMPONENT_NAME, "mTest_HeaderDictionary")

   mSeedTestData

   Dim oCWks As New f_C_Wks
   oCWks.Construct DEV_a_wks_TestCanvas

   Dim oRng As Range
   Dim oRngEndColumn As Range
   Set oRng = oCWks.oWks_prop_r.Range("B3")
   Set oRngEndColumn = oCWks.oWks_prop_r.Range("D3")

   oCWks.SetDataRangeByAnchors oRng, oRngEndColumn
   oCWks.CreateHeaderDictionary 1

   oC_Test.oC_prop_r_Assert.AssertEqual 4, oCWks.l_prop_r_ColumnNumberByHeaderName("Test-$D$1"), _
      "Header 'Test-$D$1' should map to column 4"

   DEV_f_p_CompleteTest oC_Test
End Sub

' Purpose: tests that SanitizeUsedRange removes excess used range
' 0.3.0    19.03.2026    Claude Code    Migrated from Debug.Assert to assertion framework
Public Sub mTest_SanitizeUsedRange()
   Dim oC_Test As DEV_f_C_UnitTest
   Set oC_Test = oC_DEV_f_p_CreateTest( _
      "SanitizeUsedRange trims used range to actual data extent", _
      s_m_COMPONENT_NAME, "mTest_SanitizeUsedRange")

   mSeedTestData

   Dim oCWks As New f_C_Wks
   oCWks.Construct DEV_a_wks_TestCanvas

   oCWks.SanitizeUsedRange
   oC_Test.oC_prop_r_Assert.AssertEqual "$A$1:$E$7", oCWks.oWks_prop_r.UsedRange.Address, _
      "UsedRange after sanitization should be $A$1:$E$7"

   DEV_f_p_CompleteTest oC_Test
End Sub

' Purpose: tests that SetDataRangeByAnchors with header flag creates correct header dict
' 0.3.0    19.03.2026    Claude Code    Migrated from Debug.Assert to assertion framework
Public Sub mTest_SetDataRangeWithHeaderDict()
   Dim oC_Test As DEV_f_C_UnitTest
   Set oC_Test = oC_DEV_f_p_CreateTest( _
      "SetDataRangeByAnchors with bFirstRowIsHeader creates correct header dictionary", _
      s_m_COMPONENT_NAME, "mTest_SetDataRangeWithHeaderDict")

   mSeedTestData

   Dim oCWks As New f_C_Wks
   oCWks.Construct DEV_a_wks_TestCanvas

   Dim oRng As Range
   Dim oRngEndColumn As Range
   Set oRng = oCWks.oWks_prop_r.Range("B3")
   Set oRngEndColumn = oCWks.oWks_prop_r.Range("D3")

   oCWks.SetDataRangeByAnchors oRng, oRngEndColumn, True, True
   oC_Test.oC_prop_r_Assert.AssertEqual 3, oCWks.l_prop_r_ColumnNumberByHeaderName("Test-$C$3"), _
      "Header 'Test-$C$3' should map to column 3"

   DEV_f_p_CompleteTest oC_Test
End Sub

' Purpose: tests that worksheet change events are correctly fired/suppressed
' 0.3.0    19.03.2026    Claude Code    Migrated from Debug.Assert to assertion framework
Public Sub mTest_WksChangeEvent()
   Dim oC_Test As DEV_f_C_UnitTest
   Set oC_Test = oC_DEV_f_p_CreateTest( _
      "Worksheet change event fires only when active", _
      s_m_COMPONENT_NAME, "mTest_WksChangeEvent")

   mSeedTestData

   Dim oCWks As New f_C_Wks
   oCWks.Construct DEV_a_wks_TestCanvas

   b_m_ChangeEventInvoked = False
   oCWks.s_prop_rw_NameOfSubToRunOnWksChange = "mOnChangeTest"

   oCWks.oWks_prop_r.Range("$L$1").Value2 = "Change!"
   oC_Test.oC_prop_r_Assert.AssertFalse b_m_ChangeEventInvoked, _
      "Change event should not fire when event is inactive"

   oCWks.b_prop_rw_WksChangeEventIsActive = True
   oCWks.oWks_prop_r.Range("$L$1").Value2 = "Change!"
   oC_Test.oC_prop_r_Assert.AssertTrue b_m_ChangeEventInvoked, _
      "Change event should fire when event is active"

   DEV_f_p_CompleteTest oC_Test
End Sub

' --- Private Helpers ---

' Purpose: seeds test data on the test canvas for reuse across tests
' 0.3.0    19.03.2026    Claude Code    Extracted from mTest_f_C_Wks
Private Sub mSeedTestData()
   Dim oCWks As New f_C_Wks
   Dim oRngCell As Range

   oCWks.Construct DEV_a_wks_TestCanvas
   oCWks.DeleteAllContents
   For Each oRngCell In oCWks.oWks_prop_r.Range("A1:E7")
      oRngCell.Value2 = "Test-" & oRngCell.Address
   Next oRngCell
   oCWks.oWks_prop_r.Range("B8").Interior.Color = vbWhite
End Sub

Private Sub mOnChangeTest(ByRef oRng_arg_Target As Range)
   b_m_ChangeEventInvoked = True
End Sub
