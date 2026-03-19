Attribute VB_Name = "DEV_f_pM_Testing"
' CORE-DEV - do not change, optionally remove when deploying app
'============================================================================================
'   NAME:     DEV_f_pM_Testing
'============================================================================================
'   Purpose:  Core unit testing orchestration: registration, execution, and coordination
'             of the unit test framework
'   Access:   Private
'   Type:     Module
'   Author:   Guenther Lehner
'   Contact:  guenther.lehner@protonmail.com
'   GitHubID: gueleh
'   Required: DEV_f_C_UnitTest, DEV_f_C_Assert, DEV_f_C_AssertResult,
'             DEV_f_C_TestReport, DEV_f_pM_TestRegistry
'   Usage:    Call DEV_f_p_RunAllTests to execute all registered tests.
'             Use DEV_f_p_CreateTest / DEV_f_p_CompleteTest in test procedures.
'--------------------------------------------------------------------------------------------
'   VERSION HISTORY
'   Version    Date    Developer    Changes
'   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   0.2.0    19.03.2026    Claude Code    Added test runner, CreateTest/CompleteTest API,
'         fixed RegisterExecutionError bug
'   0.10.0    04.08.2022    gueleh    Changed property names to meet new convention
'   0.1.0    20220709    gueleh    Initially created
'--------------------------------------------------------------------------------------------
'   BACKLOG
'   ''''''''''''''''''''
'   none
'============================================================================================
Option Explicit
Option Private Module

Private Const s_m_COMPONENT_NAME As String = "DEV_f_pM_Testing"

' Purpose: registers a unit test for later execution and evaluation when supposed to be tested
' 0.1.0    20220709    gueleh    Initially created
Public Sub DEV_f_p_RegisterUnitTest _
( _
   ByRef oC_arg_CallParams As f_C_CallParams _
)

   If oC_f_p_FrameworkSettings.b_prop_rw_ThisIsATestRun Then
      Dim oCUnitTest As New DEV_f_C_UnitTest
      oC_arg_CallParams.l_prop_rw_UnitTestIndex = oCol_f_p_UnitTests.Count + 1
      oCUnitTest.InitializeUnitTest oC_arg_CallParams
      oCol_f_p_UnitTests.Add oCUnitTest
   End If
End Sub

' Purpose: registers an execution error when tested
' 0.2.0    19.03.2026    Claude Code    Fixed bug: condition was = 0, must be > 0;
'         now actually calls RegisterExecutionError on the UnitTest object
' 0.1.0    20220709    gueleh    Initially created
Public Sub DEV_f_p_RegisterExecutionError _
( _
   ByRef oC_arg_CallParams As f_C_CallParams _
)

   If oC_f_p_FrameworkSettings.b_prop_rw_ThisIsATestRun And oC_arg_CallParams.l_prop_rw_UnitTestIndex > 0 Then
      Dim oCUnitTest As DEV_f_C_UnitTest
      Set oCUnitTest = oCol_f_p_UnitTests(oC_arg_CallParams.l_prop_rw_UnitTestIndex)
      oCUnitTest.RegisterExecutionError oC_arg_CallParams
   End If
End Sub

' Purpose: creates a new unit test with description, for use in test procedures.
'         The test is added to oCol_f_p_UnitTests and returned for assertion access.
' 0.2.0    19.03.2026    Claude Code    Initially created
Public Function oC_DEV_f_p_CreateTest _
( _
   ByVal s_arg_TestDescription As String, _
   ByVal s_arg_ComponentName As String, _
   ByVal s_arg_ProcedureName As String _
) As DEV_f_C_UnitTest

   Dim oCTest As New DEV_f_C_UnitTest
   oCTest.InitializeWithDescription s_arg_TestDescription, s_arg_ComponentName, s_arg_ProcedureName
   oCol_f_p_UnitTests.Add oCTest
   Set oC_DEV_f_p_CreateTest = oCTest
End Function

' Purpose: marks a test as completed after all assertions have been made
' 0.2.0    19.03.2026    Claude Code    Initially created
Public Sub DEV_f_p_CompleteTest _
( _
   ByRef oC_arg_Test As DEV_f_C_UnitTest _
)

   oC_arg_Test.MarkCompleted
End Sub

' Purpose: runs all registered tests and outputs the report
' 0.2.0    19.03.2026    Claude Code    Initially created
Public Sub DEV_f_p_RunAllTests()
   Dim dteStart As Date
   dteStart = Now()

   f_p_InitGlobals
   oC_f_p_FrameworkSettings.b_prop_rw_ThisIsATestRun = True

   DEV_f_p_RegisterAllTests
   DEV_f_p_ExecuteAllTests

   Dim oC_Report As New DEV_f_C_TestReport
   oC_Report.GenerateReport oCol_f_p_UnitTests, dteStart

   oC_f_p_FrameworkSettings.b_prop_rw_ThisIsATestRun = False
End Sub

' Purpose: runs all registered tests and outputs results to worksheet
' 0.2.0    19.03.2026    Claude Code    Initially created
Public Sub DEV_f_p_RunAllTestsToWorksheet()
   Dim dteStart As Date
   dteStart = Now()

   f_p_InitGlobals
   oC_f_p_FrameworkSettings.b_prop_rw_ThisIsATestRun = True

   DEV_f_p_RegisterAllTests
   DEV_f_p_ExecuteAllTests

   Dim oC_Report As New DEV_f_C_TestReport
   oC_Report.GenerateReportToWorksheet oCol_f_p_UnitTests, dteStart, DEV_f_wks_TestCanvas

   oC_f_p_FrameworkSettings.b_prop_rw_ThisIsATestRun = False
End Sub
