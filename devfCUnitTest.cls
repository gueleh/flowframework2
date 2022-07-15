VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "devfCUnitTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' -------------------------------------------------------------------------------------------
' CORE-DEV, do not change
'============================================================================================
'   NAME:     devfCUnitTest
'============================================================================================
'   Purpose:  a unit test
'   Access:   Public
'   Type:     Class Module
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

Private Const smCOMPONENT_NAME As String = "devfCUnitTest"

Private bmExecutionError As Boolean 'indicates whether the unit has executed successfully
Private omCCallParams As fCCallParams 'stores the call parameters of the tested unit

' Purpose: initializes class instance for a unit test
' 0.1.0    20220709    gueleh    Initially created
Public Sub InitializeUnitTest(ByRef oCCallParams As fCCallParams)
   Set omCCallParams = oCCallParams
End Sub

' Purpose: registers an execution error for a unit test
' 0.1.0    20220709    gueleh    Initially created
Public Sub RegisterExecutionError(ByRef oCCallParams As fCCallParams)
   bmExecutionError = True
   omCCallParams.SetoCError oCCallParams.oCError
End Sub
