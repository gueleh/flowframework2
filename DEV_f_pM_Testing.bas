Attribute VB_Name = "DEV_f_pM_Testing"
' CORE-DEV - do not change, optionally remove when deploying app
'============================================================================================
'   NAME:     DEV_f_pM_Testing
'============================================================================================
'   Purpose:  running the unit and integration tests
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
' 0.10.0    04.08.2022    gueleh    Changed property names to meet new convention
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
' 0.1.0    20220709    gueleh    Initially created
Public Sub DEV_f_p_RegisterExecutionError _
( _
   ByRef oC_arg_CallParams As f_C_CallParams _
)
   
   If oC_f_p_FrameworkSettings.b_prop_rw_ThisIsATestRun And oC_arg_CallParams.l_prop_rw_UnitTestIndex = 0 Then
      Dim oCUnitTest As DEV_f_C_UnitTest
      Set oCUnitTest = oCol_f_p_UnitTests(oC_arg_CallParams.l_prop_rw_UnitTestIndex)
   End If
End Sub

' Purpose: runs the existing unit tests
' 0.1.0    20220709    gueleh    Initially created
Private Sub DEV_f_m_RunUnitTests()
   f_p_InitGlobals
   oC_f_p_FrameworkSettings.b_prop_rw_ThisIsATestRun = True
   'TODO: DEV_f_m_RunUnitTests: add the actual code for running the unit tests
End Sub
