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
'   0.1.0    20220709    gueleh    Initially created
'--------------------------------------------------------------------------------------------
'   BACKLOG
'   ''''''''''''''''''''
'   none
'============================================================================================
Option Explicit
Option Private Module

Private Const smCOMPONENT_NAME As String = "devfpMTesting"

' Purpose: registers a unit test for later execution and evaluation when supposed to be tested
' 0.1.0    20220709    gueleh    Initially created
Public Sub devfRegisterUnitTest(ByRef oCCallParams As fCCallParams)
   If oCfgFrameworkSettings.bThisIsATestRun Then
      Dim oCUnitTest As New devfCUnitTest
      oCCallParams.lUnitTestIndex = oColfgUnitTests.Count + 1
      oCUnitTest.InitializeUnitTest oCCallParams
      oColfgUnitTests.Add oCUnitTest
   End If
End Sub

' Purpose: registers an execution error when tested
' 0.1.0    20220709    gueleh    Initially created
Public Sub devfRegisterExecutionError(ByRef oCCallParams As fCCallParams)
   If oCfgFrameworkSettings.bThisIsATestRun And oCCallParams.lUnitTestIndex = 0 Then
      Dim oCUnitTest As devfCUnitTest
      Set oCUnitTest = oColfgUnitTests(oCCallParams.lUnitTestIndex)
   End If
End Sub

' Purpose: runs the existing unit tests
' 0.1.0    20220709    gueleh    Initially created
Private Sub mRunUnitTests()
   fInitGlobals
   oCfgFrameworkSettings.bThisIsATestRun = True
End Sub
