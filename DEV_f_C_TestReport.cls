VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DEV_f_C_TestReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' CORE-DEV - do not change, optionally remove when deploying app
'============================================================================================
'   NAME:     DEV_f_C_TestReport
'============================================================================================
'   Purpose:  Generates test result reports from the unit test collection.
'             Supports output to Immediate Window (Debug.Print) and to a worksheet.
'   Access:   Public
'   Type:     Class Module
'   Author:   Claude Code
'   Contact:  guenther.lehner@protonmail.com
'   GitHubID: gueleh
'   Required: DEV_f_C_UnitTest, DEV_f_C_Assert, DEV_f_C_AssertResult
'   Usage:    Dim oC_Report As New DEV_f_C_TestReport
'             oC_Report.GenerateReport oCol_f_p_UnitTests, dteStartTime
'--------------------------------------------------------------------------------------------
'   VERSION HISTORY
'   Version    Date    Developer    Changes
'   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   0.1.0    19.03.2026    Claude Code    Initially created
'--------------------------------------------------------------------------------------------
'   BACKLOG
'   ''''''''''''''''''''
'   none
'============================================================================================
Option Explicit

Private Const s_m_COMPONENT_NAME As String = "DEV_f_C_TestReport"

Private Const s_m_SEPARATOR As String = "--------------------------------------------------------------------------------"
Private Const s_m_HEADER_SEPARATOR As String = "================================================================================"

' Purpose: generates a test report to the Immediate Window (Debug.Print)
' 0.1.0    19.03.2026    Claude Code    Initially created
Public Sub GenerateReport _
( _
   ByRef oCol_arg_Tests As Collection, _
   ByVal dte_arg_StartTime As Date _
)

   Dim lTotalTests As Long
   Dim lPassedTests As Long
   Dim lFailedTests As Long
   Dim lTotalAssertions As Long
   Dim lPassedAssertions As Long
   Dim lFailedAssertions As Long
   Dim dblTotalDuration As Double
   Dim oCTest As DEV_f_C_UnitTest
   Dim oCResult As DEV_f_C_AssertResult
   Dim lIndex As Long

   If oCol_arg_Tests Is Nothing Then
      Debug.Print "No test collection available."
      Exit Sub
   End If

   lTotalTests = oCol_arg_Tests.Count
   If lTotalTests = 0 Then
      Debug.Print "No tests were executed."
      Exit Sub
   End If

   dblTotalDuration = CDbl(Now() - dte_arg_StartTime) * 86400

   Debug.Print vbNewLine & s_m_HEADER_SEPARATOR
   Debug.Print "  UNIT TEST REPORT - " & Format(Now(), "YYYY-MM-DD hh:mm:ss")
   Debug.Print s_m_HEADER_SEPARATOR

   For lIndex = 1 To oCol_arg_Tests.Count
      Set oCTest = oCol_arg_Tests(lIndex)

      If oCTest.b_prop_r_Passed Then
         lPassedTests = lPassedTests + 1
      Else
         lFailedTests = lFailedTests + 1
      End If

      Debug.Print vbNewLine & oCTest.s_prop_r_Summary()

      lTotalAssertions = lTotalAssertions + oCTest.oC_prop_r_Assert.l_prop_r_TotalCount
      lPassedAssertions = lPassedAssertions + oCTest.oC_prop_r_Assert.l_prop_r_PassedCount
      lFailedAssertions = lFailedAssertions + oCTest.oC_prop_r_Assert.l_prop_r_FailedCount

      If oCTest.oC_prop_r_Assert.l_prop_r_FailedCount > 0 Then
         Dim lAssertIndex As Long
         For lAssertIndex = 1 To oCTest.oC_prop_r_Assert.oCol_prop_r_Results.Count
            Set oCResult = oCTest.oC_prop_r_Assert.oCol_prop_r_Results(lAssertIndex)
            If Not oCResult.b_prop_r_Passed Then
               Debug.Print "   " & oCResult.s_prop_r_Summary()
            End If
         Next lAssertIndex
      End If
   Next lIndex

   Debug.Print vbNewLine & s_m_SEPARATOR
   Debug.Print "  SUMMARY"
   Debug.Print s_m_SEPARATOR
   Debug.Print "  Tests:      " & lTotalTests & " total, " & lPassedTests & " passed, " & lFailedTests & " failed"
   Debug.Print "  Assertions: " & lTotalAssertions & " total, " & lPassedAssertions & " passed, " & lFailedAssertions & " failed"
   Debug.Print "  Duration:   " & Format(dblTotalDuration, "0.00") & " seconds"

   If lFailedTests = 0 Then
      Debug.Print vbNewLine & "  ALL TESTS PASSED"
   Else
      Debug.Print vbNewLine & "  " & lFailedTests & " TEST(S) FAILED"
   End If

   Debug.Print s_m_HEADER_SEPARATOR
End Sub

' Purpose: generates a test report to a worksheet
' 0.1.0    19.03.2026    Claude Code    Initially created
Public Sub GenerateReportToWorksheet _
( _
   ByRef oCol_arg_Tests As Collection, _
   ByVal dte_arg_StartTime As Date, _
   ByRef oWks_arg_Target As Worksheet _
)

   Dim lRow As Long
   Dim lTotalTests As Long
   Dim lPassedTests As Long
   Dim lFailedTests As Long
   Dim lTotalAssertions As Long
   Dim lPassedAssertions As Long
   Dim lFailedAssertions As Long
   Dim dblTotalDuration As Double
   Dim oCTest As DEV_f_C_UnitTest
   Dim oCResult As DEV_f_C_AssertResult
   Dim lIndex As Long
   Dim lAssertIndex As Long

   If oCol_arg_Tests Is Nothing Or oWks_arg_Target Is Nothing Then Exit Sub

   oWks_arg_Target.Cells.Delete
   dblTotalDuration = CDbl(Now() - dte_arg_StartTime) * 86400

   lRow = 1
   With oWks_arg_Target
      .Cells(lRow, 1).Value2 = "UNIT TEST REPORT"
      .Cells(lRow, 2).Value2 = Format(Now(), "YYYY-MM-DD hh:mm:ss")
      .Cells(lRow, 1).Font.Bold = True

      lRow = 3
      .Cells(lRow, 1).Value2 = "Component"
      .Cells(lRow, 2).Value2 = "Procedure"
      .Cells(lRow, 3).Value2 = "Description"
      .Cells(lRow, 4).Value2 = "Result"
      .Cells(lRow, 5).Value2 = "Assertions"
      .Cells(lRow, 6).Value2 = "Passed"
      .Cells(lRow, 7).Value2 = "Failed"
      .Cells(lRow, 8).Value2 = "Duration (s)"
      .Cells(lRow, 9).Value2 = "Details"
      .Range(.Cells(lRow, 1), .Cells(lRow, 9)).Font.Bold = True

      lRow = 4
      For lIndex = 1 To oCol_arg_Tests.Count
         Set oCTest = oCol_arg_Tests(lIndex)

         If oCTest.b_prop_r_Passed Then
            lPassedTests = lPassedTests + 1
         Else
            lFailedTests = lFailedTests + 1
         End If

         lTotalAssertions = lTotalAssertions + oCTest.oC_prop_r_Assert.l_prop_r_TotalCount
         lPassedAssertions = lPassedAssertions + oCTest.oC_prop_r_Assert.l_prop_r_PassedCount
         lFailedAssertions = lFailedAssertions + oCTest.oC_prop_r_Assert.l_prop_r_FailedCount

         .Cells(lRow, 1).Value2 = oCTest.s_prop_r_ComponentName
         .Cells(lRow, 2).Value2 = oCTest.s_prop_r_ProcedureName
         .Cells(lRow, 3).Value2 = oCTest.s_prop_r_TestDescription

         If oCTest.b_prop_r_Passed Then
            .Cells(lRow, 4).Value2 = "PASS"
         Else
            .Cells(lRow, 4).Value2 = "FAIL"
         End If

         .Cells(lRow, 5).Value2 = oCTest.oC_prop_r_Assert.l_prop_r_TotalCount
         .Cells(lRow, 6).Value2 = oCTest.oC_prop_r_Assert.l_prop_r_PassedCount
         .Cells(lRow, 7).Value2 = oCTest.oC_prop_r_Assert.l_prop_r_FailedCount

         If oCTest.b_prop_r_Completed Then
            .Cells(lRow, 8).Value2 = Format(oCTest.dbl_prop_r_DurationSeconds, "0.000")
         End If

         If oCTest.b_prop_r_ExecutionError Then
            .Cells(lRow, 9).Value2 = "EXECUTION ERROR"
         ElseIf oCTest.oC_prop_r_Assert.l_prop_r_FailedCount > 0 Then
            Dim sDetails As String
            sDetails = vbNullString
            For lAssertIndex = 1 To oCTest.oC_prop_r_Assert.oCol_prop_r_Results.Count
               Set oCResult = oCTest.oC_prop_r_Assert.oCol_prop_r_Results(lAssertIndex)
               If Not oCResult.b_prop_r_Passed Then
                  If Len(sDetails) > 0 Then sDetails = sDetails & "; "
                  sDetails = sDetails & oCResult.s_prop_r_Summary()
               End If
            Next lAssertIndex
            .Cells(lRow, 9).Value2 = sDetails
         End If

         lRow = lRow + 1
      Next lIndex

      lTotalTests = oCol_arg_Tests.Count
      lRow = lRow + 1
      .Cells(lRow, 1).Value2 = "SUMMARY"
      .Cells(lRow, 1).Font.Bold = True
      lRow = lRow + 1
      .Cells(lRow, 1).Value2 = "Tests:"
      .Cells(lRow, 2).Value2 = lTotalTests & " total, " & lPassedTests & " passed, " & lFailedTests & " failed"
      lRow = lRow + 1
      .Cells(lRow, 1).Value2 = "Assertions:"
      .Cells(lRow, 2).Value2 = lTotalAssertions & " total, " & lPassedAssertions & " passed, " & lFailedAssertions & " failed"
      lRow = lRow + 1
      .Cells(lRow, 1).Value2 = "Duration:"
      .Cells(lRow, 2).Value2 = Format(dblTotalDuration, "0.00") & " seconds"

      .Columns.AutoFit
   End With

   GenerateReport oCol_arg_Tests, dte_arg_StartTime
End Sub
