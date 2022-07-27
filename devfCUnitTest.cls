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

Private Const s_m_COMPONENT_NAME As String = "devfCUnitTest"

Private b_m_ExecutionError As Boolean 'indicates whether the unit has executed successfully
Private oC_m_CallParams As fCCallParams 'stores the call parameters of the tested unit

' Purpose: initializes class instance for a unit test
' 0.1.0    20220709    gueleh    Initially created
Public Sub InitializeUnitTest(ByRef oC_CallParams As fCCallParams)
   Set oC_m_CallParams = oC_CallParams
End Sub

' Purpose: registers an execution error for a unit test
' 0.1.0    20220709    gueleh    Initially created
Public Sub RegisterExecutionError(ByRef oC_CallParams As fCCallParams)
   b_m_ExecutionError = True
   oC_m_CallParams.SetoCError oC_CallParams.oCError
End Sub
