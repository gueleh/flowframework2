Attribute VB_Name = "devfmTesting"
' -------------------------------------------------------------------------------------------
' CORE-DEV, do not change
'============================================================================================
'   NAME:     devfmTesting
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

Private Const msCOMPONENT_NAME As String = "devfmTesting"

' Purpose: registers a unit test for later execution and evaluation when supposed to be tested
' 0.1.0    20220709    gueleh    Initially created
Public Sub devfRegisterUnitTest(ByRef clsCallParams As fclsCallParams)
   If fgclsFrameworkSettings.bThisIsATestRun Then
      Dim clsUnitTest As New devfclsUnitTest
      clsCallParams.lUnitTestIndex = fgcolUnitTests.Count + 1
      clsUnitTest.InitializeUnitTest clsCallParams
      fgcolUnitTests.Add clsUnitTest
   End If
End Sub

' Purpose: registers an execution error when tested
' 0.1.0    20220709    gueleh    Initially created
Public Sub devfRegisterExecutionError(ByRef clsCallParams As fclsCallParams)
   If fgclsFrameworkSettings.bThisIsATestRun And clsCallParams.lUnitTestIndex = 0 Then
      Dim clsUnitTest As devfclsUnitTest
      Set clsUnitTest = fgcolUnitTests(clsCallParams.lUnitTestIndex)
   End If
End Sub

' Purpose: runs the existing unit tests
' 0.1.0    20220709    gueleh    Initially created
Private Sub mRunUnitTests()
   fInitGlobals
   fgclsFrameworkSettings.bThisIsATestRun = True
End Sub
