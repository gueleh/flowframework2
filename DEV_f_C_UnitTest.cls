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
'   Purpose:  Represents a single unit test with assertions, timing, and result tracking
'   Access:   Public
'   Type:     Class Module
'   Author:   Guenther Lehner
'   Contact:  guenther.lehner@protonmail.com
'   GitHubID: gueleh
'   Required: DEV_f_C_Assert, DEV_f_C_AssertResult, f_C_CallParams
'   Usage:    Created by DEV_f_p_CreateTest, assertions via oC_prop_r_Assert,
'             completed by DEV_f_p_CompleteTest
'--------------------------------------------------------------------------------------------
'   VERSION HISTORY
'   Version    Date    Developer    Changes
'   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   0.2.0    19.03.2026    Claude Code    Extended with Assert, description, timing, result tracking
'   0.1.0    20220709    gueleh    Initially created
'--------------------------------------------------------------------------------------------
'   BACKLOG
'   ''''''''''''''''''''
'   none
'============================================================================================
Option Explicit

Private Const s_m_COMPONENT_NAME As String = "DEV_f_C_UnitTest"

Private b_m_ExecutionError As Boolean
Private oC_m_CallParams As f_C_CallParams
Private oC_m_Assert As DEV_f_C_Assert
Private s_m_TestDescription As String
Private s_m_ComponentName As String
Private s_m_ProcedureName As String
Private dte_m_StartTime As Date
Private dte_m_EndTime As Date
Private b_m_Completed As Boolean

' --- Properties ---

Public Property Get b_prop_r_ExecutionError() As Boolean: b_prop_r_ExecutionError = b_m_ExecutionError: End Property

Public Property Get oC_prop_r_CallParams() As f_C_CallParams: Set oC_prop_r_CallParams = oC_m_CallParams: End Property

Public Property Get oC_prop_r_Assert() As DEV_f_C_Assert: Set oC_prop_r_Assert = oC_m_Assert: End Property

Public Property Get s_prop_r_TestDescription() As String: s_prop_r_TestDescription = s_m_TestDescription: End Property

Public Property Get s_prop_r_ComponentName() As String: s_prop_r_ComponentName = s_m_ComponentName: End Property

Public Property Get s_prop_r_ProcedureName() As String: s_prop_r_ProcedureName = s_m_ProcedureName: End Property

Public Property Get dte_prop_r_StartTime() As Date: dte_prop_r_StartTime = dte_m_StartTime: End Property

Public Property Get dte_prop_r_EndTime() As Date: dte_prop_r_EndTime = dte_m_EndTime: End Property

Public Property Get b_prop_r_Completed() As Boolean: b_prop_r_Completed = b_m_Completed: End Property

' Purpose: returns True if the test passed (no execution error and all assertions passed)
' 0.2.0    19.03.2026    Claude Code    Initially created
Public Property Get b_prop_r_Passed() As Boolean
   If b_m_ExecutionError Then
      b_prop_r_Passed = False
   ElseIf oC_m_Assert.l_prop_r_TotalCount = 0 Then
      b_prop_r_Passed = Not b_m_ExecutionError
   Else
      b_prop_r_Passed = oC_m_Assert.b_prop_r_AllPassed
   End If
End Property

' Purpose: returns the duration of the test in seconds
' 0.2.0    19.03.2026    Claude Code    Initially created
Public Property Get dbl_prop_r_DurationSeconds() As Double
   If b_m_Completed Then
      dbl_prop_r_DurationSeconds = CDbl(dte_m_EndTime - dte_m_StartTime) * 86400
   Else
      dbl_prop_r_DurationSeconds = CDbl(Now() - dte_m_StartTime) * 86400
   End If
End Property

' --- Lifecycle ---

Private Sub Class_Initialize()
   Set oC_m_Assert = New DEV_f_C_Assert
   dte_m_StartTime = Now()
End Sub

' Purpose: initializes class instance for a unit test (legacy interface, kept for compatibility)
' 0.1.0    20220709    gueleh    Initially created
Public Sub InitializeUnitTest _
( _
   ByRef oC_arg_CallParams As f_C_CallParams _
)

   Set oC_m_CallParams = oC_arg_CallParams
   s_m_ComponentName = oC_arg_CallParams.s_prop_rw_ComponentName
   s_m_ProcedureName = oC_arg_CallParams.s_prop_rw_ProcedureName
End Sub

' Purpose: initializes the test with a description for the new test API
' 0.2.0    19.03.2026    Claude Code    Initially created
Public Sub InitializeWithDescription _
( _
   ByVal s_arg_TestDescription As String, _
   ByVal s_arg_ComponentName As String, _
   ByVal s_arg_ProcedureName As String _
)

   s_m_TestDescription = s_arg_TestDescription
   s_m_ComponentName = s_arg_ComponentName
   s_m_ProcedureName = s_arg_ProcedureName
End Sub

' Purpose: registers an execution error for a unit test
' 0.1.0    20220709    gueleh    Initially created
Public Sub RegisterExecutionError _
( _
   ByRef oC_arg_CallParams As f_C_CallParams _
)

   b_m_ExecutionError = True
   If Not oC_arg_CallParams.oC_prop_r_Error Is Nothing Then
      oC_m_CallParams.SetoCError oC_arg_CallParams.oC_prop_r_Error
   End If
End Sub

' Purpose: marks the test as completed and records the end time
' 0.2.0    19.03.2026    Claude Code    Initially created
Public Sub MarkCompleted()
   dte_m_EndTime = Now()
   b_m_Completed = True
End Sub

' Purpose: returns a formatted single-line summary of the test result
' 0.2.0    19.03.2026    Claude Code    Initially created
Public Function s_prop_r_Summary() As String
   Dim sResult As String
   If b_prop_r_Passed Then
      sResult = "[PASS] "
   Else
      sResult = "[FAIL] "
   End If
   sResult = sResult & s_m_ComponentName & "." & s_m_ProcedureName
   If Len(s_m_TestDescription) > 0 Then
      sResult = sResult & " - " & s_m_TestDescription
   End If
   If oC_m_Assert.l_prop_r_TotalCount > 0 Then
      sResult = sResult & " (" & oC_m_Assert.l_prop_r_PassedCount & "/" & oC_m_Assert.l_prop_r_TotalCount & " assertions)"
   End If
   If b_m_ExecutionError Then
      sResult = sResult & " [EXECUTION ERROR]"
   End If
   s_prop_r_Summary = sResult
End Function


