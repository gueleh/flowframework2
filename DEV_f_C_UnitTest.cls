VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DEV_f_C_UnitTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' CORE-DEV - do not change, optionally remove when deploying app
'============================================================================================
'   NAME:     DEV_f_C_UnitTest
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

Private Const s_m_COMPONENT_NAME As String = "DEV_f_C_UnitTest"

Private b_m_ExecutionError As Boolean 'indicates whether the unit has executed successfully
Private oC_m_CallParams As f_C_CallParams 'stores the call parameters of the tested unit

' Purpose: initializes class instance for a unit test
' 0.1.0    20220709    gueleh    Initially created
Public Sub InitializeUnitTest _
( _
   ByRef oC_arg_CallParams As f_C_CallParams _
)
   
   Set oC_m_CallParams = oC_arg_CallParams
End Sub

' Purpose: registers an execution error for a unit test
' 0.1.0    20220709    gueleh    Initially created
Public Sub RegisterExecutionError _
( _
   ByRef oC_arg_CallParams As f_C_CallParams _
)
   
   b_m_ExecutionError = True
   oC_m_CallParams.SetoCError oC_arg_CallParams.oC_prop_r_Error
End Sub
