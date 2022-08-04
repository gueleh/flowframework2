Attribute VB_Name = "devfpMTesting"
' -------------------------------------------------------------------------------------------
' CORE-DEV, do not change
'============================================================================================
'   NAME:     devfpMTesting
'============================================================================================
'   Purpose:  running the unit and integration tests
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
' 0.10.0    04.08.2022    gueleh    Changed property names to meet new convention
'   0.1.0    20220709    gueleh    Initially created
'--------------------------------------------------------------------------------------------
'   BACKLOG
'   ''''''''''''''''''''
'   none
'============================================================================================
Option Explicit
Option Private Module

Private Const s_m_COMPONENT_NAME As String = "devfpMTesting"

' Purpose: registers a unit test for later execution and evaluation when supposed to be tested
' 0.1.0    20220709    gueleh    Initially created
Public Sub DEV_f_p_RegisterUnitTest(ByRef oCCallParams As fCCallParams)
   If oC_f_p_FrameworkSettings.b_prop_rw_ThisIsATestRun Then
      Dim oCUnitTest As New devfCUnitTest
      oCCallParams.l_prop_rw_UnitTestIndex = oCol_f_p_UnitTests.Count + 1
      oCUnitTest.InitializeUnitTest oCCallParams
      oCol_f_p_UnitTests.Add oCUnitTest
   End If
End Sub

' Purpose: registers an execution error when tested
' 0.1.0    20220709    gueleh    Initially created
Public Sub DEV_f_p_RegisterExecutionError(ByRef oCCallParams As fCCallParams)
   If oC_f_p_FrameworkSettings.b_prop_rw_ThisIsATestRun And oCCallParams.l_prop_rw_UnitTestIndex = 0 Then
      Dim oCUnitTest As devfCUnitTest
      Set oCUnitTest = oCol_f_p_UnitTests(oCCallParams.l_prop_rw_UnitTestIndex)
   End If
End Sub

' Purpose: runs the existing unit tests
' 0.1.0    20220709    gueleh    Initially created
Private Sub DEV_f_m_RunUnitTests()
   f_p_InitGlobals
   oC_f_p_FrameworkSettings.b_prop_rw_ThisIsATestRun = True
   'TODO: DEV_f_m_RunUnitTests: add the actual code for running the unit tests
End Sub
