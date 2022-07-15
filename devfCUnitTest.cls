VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "devfclsUnitTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' -------------------------------------------------------------------------------------------
' CORE-DEV, do not change
'============================================================================================
'   NAME:     devfclsUnitTest
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

Private Const msCOMPONENT_NAME As String = "devfclsUnitTest"

Private mbExecutionError As Boolean
Private mclsCallParams As fclsCallParams

' Purpose: initializes class for a unit test
' 0.1.0    20220709    gueleh    Initially created
Public Sub InitializeUnitTest(ByRef clsCallParams As fclsCallParams)
   Set mclsCallParams = clsCallParams
End Sub

' Purpose: registers an execution error for a unit test
' 0.1.0    20220709    gueleh    Initially created
Public Sub RegisterExecutionError(ByRef clsCallParams As fclsCallParams)
   mbExecutionError = True
   mclsCallParams.SetError clsCallParams.clsError
End Sub
