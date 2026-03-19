Attribute VB_Name = "DEV_f_pM_TestRegistry"
' CORE-DEV - do not change, optionally remove when deploying app
'============================================================================================
'   NAME:     DEV_f_pM_TestRegistry
'============================================================================================
'   Purpose:  Central registry for all test subroutines. Each test module registers its
'             test subs here. The test runner calls DEV_f_p_RegisterAllTests to populate
'             the registry, then DEV_f_p_ExecuteAllTests to run them.
'   Access:   Private
'   Type:     Module
'   Author:   Claude Code
'   Contact:  guenther.lehner@protonmail.com
'   GitHubID: gueleh
'   Required: DEV_f_pM_Testing, DEV_f_C_UnitTest
'   Usage:    Add new test modules in DEV_f_p_RegisterAllTests.
'             Each test module must have a Public Sub DEV_f_p_RegisterTests_<ModuleName>.
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
Option Private Module

Private Const s_m_COMPONENT_NAME As String = "DEV_f_pM_TestRegistry"

' Module-level collection of test sub names (fully qualified: "ModuleName.SubName")
Private oCol_m_TestSubs As Collection

' Purpose: registers all test modules by calling their registration subs.
'         Add new test modules here as they are created.
'         Uses Application.Run so missing modules do not cause compile errors.
' 0.1.0    19.03.2026    Claude Code    Initially created
Public Sub DEV_f_p_RegisterAllTests()
   Set oCol_m_TestSubs = New Collection

   On Error Resume Next

   Application.Run "DEV_f_pM_Test_f_C_Wks.DEV_f_p_RegisterTests_f_C_Wks"
   Application.Run "DEV_f_pM_Test_f_C_DataRecord.DEV_f_p_RegisterTests_f_C_DataRecord"
   Application.Run "DEV_f_pM_Test_UtilitiesRanges.DEV_f_p_RegisterTests_UtilitiesRanges"

   On Error GoTo 0
End Sub

' Purpose: adds a test sub to the registry (called by test modules during registration)
' 0.1.0    19.03.2026    Claude Code    Initially created
Public Sub DEV_f_p_AddTestSub _
( _
   ByVal s_arg_FullyQualifiedName As String _
)

   oCol_m_TestSubs.Add s_arg_FullyQualifiedName
End Sub

' Purpose: executes all registered test subs via Application.Run.
'         Errors during individual tests are caught and do not stop the runner.
' 0.1.0    19.03.2026    Claude Code    Initially created
Public Sub DEV_f_p_ExecuteAllTests()
   Dim lIndex As Long
   Dim sTestSub As String

   If oCol_m_TestSubs Is Nothing Then Exit Sub

   For lIndex = 1 To oCol_m_TestSubs.Count
      sTestSub = oCol_m_TestSubs(lIndex)
      On Error Resume Next
      Application.Run sTestSub
      If Err.Number <> 0 Then
         Debug.Print "[ERROR] Failed to run test sub: " & sTestSub & " - " & Err.Description
         Err.Clear
      End If
      On Error GoTo 0
   Next lIndex
End Sub

' Purpose: returns the count of registered test subs
' 0.1.0    19.03.2026    Claude Code    Initially created
Public Function l_DEV_f_p_RegisteredTestCount() As Long
   If oCol_m_TestSubs Is Nothing Then
      l_DEV_f_p_RegisteredTestCount = 0
   Else
      l_DEV_f_p_RegisteredTestCount = oCol_m_TestSubs.Count
   End If
End Function
