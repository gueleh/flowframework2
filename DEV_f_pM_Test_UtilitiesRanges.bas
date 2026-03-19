Attribute VB_Name = "DEV_f_pM_Test_UtilitiesRanges"
' CORE-DEV - do not change, optionally remove when deploying app
'============================================================================================
'   NAME:     DEV_f_pM_Test_UtilitiesRanges
'============================================================================================
'   Purpose:  Unit tests for f_pM_UtilitiesRanges
'   Access:   Private
'   Type:     Module
'   Author:   Guenther Lehner
'   Contact:  guleh@pm.me
'   GitHubID: gueleh
'   Required: DEV_f_wks_TestCanvas, f_pM_UtilitiesRanges, DEV_f_pM_Testing, DEV_f_pM_TestRegistry
'   Usage:    Tests are registered via DEV_f_p_RegisterTests_UtilitiesRanges and run by
'             the test runner.
'--------------------------------------------------------------------------------------------
'   VERSION HISTORY
'   Version    Date    Developer    Changes
'   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   0.2.0    19.03.2026    Claude Code    Migrated to assertion framework
'   1.16    16.08.2024    gueleh    Initially created
'--------------------------------------------------------------------------------------------
'   BACKLOG
'   ''''''''''''''''''''
'   none
'============================================================================================
Option Explicit
Option Private Module

Private Const s_m_COMPONENT_NAME As String = "DEV_f_pM_Test_UtilitiesRanges"

' Purpose: registers all test subs of this module with the test registry
' 0.2.0    19.03.2026    Claude Code    Initially created
Public Sub DEV_f_p_RegisterTests_UtilitiesRanges()
   DEV_f_p_AddTestSub "DEV_f_pM_Test_UtilitiesRanges.mTest_RangeSizeEqual_NothingRanges"
   DEV_f_p_AddTestSub "DEV_f_pM_Test_UtilitiesRanges.mTest_RangeSizeEqual_DifferentSize"
   DEV_f_p_AddTestSub "DEV_f_pM_Test_UtilitiesRanges.mTest_RangeSizeEqual_SameEmpty"
   DEV_f_p_AddTestSub "DEV_f_pM_Test_UtilitiesRanges.mTest_RangeSizeEqual_DifferentContents"
   DEV_f_p_AddTestSub "DEV_f_pM_Test_UtilitiesRanges.mTest_RangeSizeEqual_SameContents"
End Sub

' Purpose: tests that Nothing ranges cause a processing error (function returns False)
' 0.2.0    19.03.2026    Claude Code    Migrated from Debug.Assert to assertion framework
Public Sub mTest_RangeSizeEqual_NothingRanges()
   Dim oC_Test As DEV_f_C_UnitTest
   Set oC_Test = oC_DEV_f_p_CreateTest( _
      "RangeSizeAndCellContentsAreEqual returns False for Nothing ranges", _
      s_m_COMPONENT_NAME, "mTest_RangeSizeEqual_NothingRanges")

   Dim oRngOne As Range
   Dim oRngTwo As Range
   Dim bResult As Boolean
   Dim sResult As String

   oC_Test.oC_prop_r_Assert.AssertFailure _
      b_f_p_RangeSizeAndCellContentsAreEqual(bResult, oRngOne, oRngTwo, sResult), _
      "Function should return False (processing error) for Nothing ranges"

   DEV_f_p_CompleteTest oC_Test
End Sub

' Purpose: tests that ranges of different size are not equal
' 0.2.0    19.03.2026    Claude Code    Migrated from Debug.Assert to assertion framework
Public Sub mTest_RangeSizeEqual_DifferentSize()
   Dim oC_Test As DEV_f_C_UnitTest
   Set oC_Test = oC_DEV_f_p_CreateTest( _
      "RangeSizeAndCellContentsAreEqual: different size ranges are not equal", _
      s_m_COMPONENT_NAME, "mTest_RangeSizeEqual_DifferentSize")

   DEV_Reset_DEV_f_wks_TestCanvas

   Dim oRngOne As Range
   Dim oRngTwo As Range
   Dim bResult As Boolean
   Dim sResult As String

   Set oRngOne = DEV_f_wks_TestCanvas.Range("A1:E1")
   Set oRngTwo = DEV_f_wks_TestCanvas.Range("A2:D2")

   oC_Test.oC_prop_r_Assert.AssertSuccess _
      b_f_p_RangeSizeAndCellContentsAreEqual(bResult, oRngOne, oRngTwo, sResult), _
      "Function should succeed (return True as processing result)"
   oC_Test.oC_prop_r_Assert.AssertFalse bResult, _
      "Ranges of different column count should not be equal"

   DEV_f_p_CompleteTest oC_Test
End Sub

' Purpose: tests that same-sized empty ranges are equal
' 0.2.0    19.03.2026    Claude Code    Migrated from Debug.Assert to assertion framework
Public Sub mTest_RangeSizeEqual_SameEmpty()
   Dim oC_Test As DEV_f_C_UnitTest
   Set oC_Test = oC_DEV_f_p_CreateTest( _
      "RangeSizeAndCellContentsAreEqual: same-sized empty ranges are equal", _
      s_m_COMPONENT_NAME, "mTest_RangeSizeEqual_SameEmpty")

   DEV_Reset_DEV_f_wks_TestCanvas

   Dim oRngOne As Range
   Dim oRngTwo As Range
   Dim bResult As Boolean
   Dim sResult As String

   Set oRngOne = DEV_f_wks_TestCanvas.Range("A1:E1")
   Set oRngTwo = DEV_f_wks_TestCanvas.Range("A2:E2")

   oC_Test.oC_prop_r_Assert.AssertSuccess _
      b_f_p_RangeSizeAndCellContentsAreEqual(bResult, oRngOne, oRngTwo, sResult), _
      "Function should succeed"
   oC_Test.oC_prop_r_Assert.AssertTrue bResult, _
      "Same-sized empty ranges should be equal"

   DEV_f_p_CompleteTest oC_Test
End Sub

' Purpose: tests that same-sized ranges with different contents are not equal
' 0.2.0    19.03.2026    Claude Code    Migrated from Debug.Assert to assertion framework
Public Sub mTest_RangeSizeEqual_DifferentContents()
   Dim oC_Test As DEV_f_C_UnitTest
   Set oC_Test = oC_DEV_f_p_CreateTest( _
      "RangeSizeAndCellContentsAreEqual: different cell contents are not equal", _
      s_m_COMPONENT_NAME, "mTest_RangeSizeEqual_DifferentContents")

   DEV_Reset_DEV_f_wks_TestCanvas
   DEV_f_wks_TestCanvas.Range("B1").Value2 = "Test"

   Dim oRngOne As Range
   Dim oRngTwo As Range
   Dim bResult As Boolean
   Dim sResult As String

   Set oRngOne = DEV_f_wks_TestCanvas.Range("A1:E1")
   Set oRngTwo = DEV_f_wks_TestCanvas.Range("A2:E2")

   oC_Test.oC_prop_r_Assert.AssertSuccess _
      b_f_p_RangeSizeAndCellContentsAreEqual(bResult, oRngOne, oRngTwo, sResult), _
      "Function should succeed"
   oC_Test.oC_prop_r_Assert.AssertFalse bResult, _
      "Ranges with different cell contents should not be equal"

   DEV_f_p_CompleteTest oC_Test
End Sub

' Purpose: tests that same-sized ranges with same contents are equal
' 0.2.0    19.03.2026    Claude Code    Migrated from Debug.Assert to assertion framework
Public Sub mTest_RangeSizeEqual_SameContents()
   Dim oC_Test As DEV_f_C_UnitTest
   Set oC_Test = oC_DEV_f_p_CreateTest( _
      "RangeSizeAndCellContentsAreEqual: same cell contents are equal", _
      s_m_COMPONENT_NAME, "mTest_RangeSizeEqual_SameContents")

   DEV_Reset_DEV_f_wks_TestCanvas
   DEV_f_wks_TestCanvas.Range("B1").Value2 = "Test"
   DEV_f_wks_TestCanvas.Range("B2").Value2 = "Test"

   Dim oRngOne As Range
   Dim oRngTwo As Range
   Dim bResult As Boolean
   Dim sResult As String

   Set oRngOne = DEV_f_wks_TestCanvas.Range("A1:E1")
   Set oRngTwo = DEV_f_wks_TestCanvas.Range("A2:E2")

   oC_Test.oC_prop_r_Assert.AssertSuccess _
      b_f_p_RangeSizeAndCellContentsAreEqual(bResult, oRngOne, oRngTwo, sResult), _
      "Function should succeed"
   oC_Test.oC_prop_r_Assert.AssertTrue bResult, _
      "Ranges with same cell contents should be equal"

   DEV_f_p_CompleteTest oC_Test
End Sub


