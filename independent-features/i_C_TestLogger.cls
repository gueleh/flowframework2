VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "i_C_TestLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'   NAME:     i_C_TestLogger
'============================================================================================
'   Purpose:  prints test results to provided wks
'   Access:   Public
'   Type:     Class Module
'   Author:   Günther Lehner
'   Contact:  guleh@pm.me
'   GitHubID: gueleh
'   Required:
'   Usage:
'--------------------------------------------------------------------------------------------
'   VERSION HISTORY
'   Version    Date    Developer    Changes
'   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   1.0.0    02.06.2023    gueleh    Initially created
'--------------------------------------------------------------------------------------------
'   BACKLOG
'   ''''''''''''''''''''
'   none
'============================================================================================
Option Explicit

Public Sub LogTestResults(ByRef oWksLog As Worksheet, ByRef oColTests As Collection)
   
   Dim lRow As Long
   Dim vaHeader() As Variant
   Dim oRngCell As Range
   Dim oCTest As i_C_Test
   Dim lPassed As Long
   
   If oColTests.Count > 0 Then
      'Prepare sheet for result logging
      If oWksLog.AutoFilterMode Then oWksLog.AutoFilterMode = False
      oWksLog.Cells.Clear
      lRow = 4
      vaHeader = Array("ID", "Description", "Input", "Expected Output", "Actual Output", "Starting State", "Passed?")
      For Each oRngCell In oWksLog.Range("A" & (lRow - 1)).Resize(1, UBound(vaHeader) + 1)
         oRngCell.Value2 = vaHeader(oRngCell.Column - 1)
      Next oRngCell
   
      'Log test results
      For Each oCTest In oColTests
         oWksLog.Range("A" & lRow).Resize(1, UBound(vaHeader) + 1).Value2 = oCTest.vaGetTestData
         If oCTest.bPassed() Then lPassed = lPassed + 1
         lRow = lRow + 1
      Next oCTest
      
      vaHeader = Array("Tests", oColTests.Count, "Passed", lPassed, "All Passed?", oColTests.Count = lPassed)
      For Each oRngCell In oWksLog.Range("a1").Resize(1, UBound(vaHeader) + 1)
         oRngCell.Value2 = vaHeader(oRngCell.Column - 1)
      Next oRngCell
      
      oWksLog.Columns.AutoFit
      oWksLog.Range("A3").AutoFilter
   End If
   
End Sub

