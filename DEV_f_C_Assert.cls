VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DEV_f_C_Assert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' CORE-DEV - do not change, optionally remove when deploying app
'============================================================================================
'   NAME:     DEV_f_C_Assert
'============================================================================================
'   Purpose:  Provides non-stopping assertion methods for unit testing. Unlike Debug.Assert,
'             failed assertions do not halt execution but are collected for later reporting.
'   Access:   Public
'   Type:     Class Module
'   Author:   Claude Code
'   Contact:  guenther.lehner@protonmail.com
'   GitHubID: gueleh
'   Required: DEV_f_C_AssertResult
'   Usage:    Each DEV_f_C_UnitTest instance owns one DEV_f_C_Assert instance.
'             Call assertion methods during the test, then read results via properties.
'--------------------------------------------------------------------------------------------
'   VERSION HISTORY
'   Version    Date    Developer    Changes
'   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   0.1.0    19.03.2026    Claude Code    Initially created
'--------------------------------------------------------------------------------------------
'   BACKLOG
'   ''''''''''''''''''''
'   none
'============================================================================================
Option Explicit

Private Const s_m_COMPONENT_NAME As String = "DEV_f_C_Assert"

Private oCol_m_Results As Collection
Private l_m_PassedCount As Long
Private l_m_FailedCount As Long

' --- Properties ---

Public Property Get l_prop_r_TotalCount() As Long: l_prop_r_TotalCount = oCol_m_Results.Count: End Property

Public Property Get l_prop_r_PassedCount() As Long: l_prop_r_PassedCount = l_m_PassedCount: End Property

Public Property Get l_prop_r_FailedCount() As Long: l_prop_r_FailedCount = l_m_FailedCount: End Property

Public Property Get b_prop_r_AllPassed() As Boolean: b_prop_r_AllPassed = (l_m_FailedCount = 0 And oCol_m_Results.Count > 0): End Property

Public Property Get oCol_prop_r_Results() As Collection: Set oCol_prop_r_Results = oCol_m_Results: End Property

' --- Lifecycle ---

Private Sub Class_Initialize()
   Set oCol_m_Results = New Collection
End Sub

' --- Assertion Methods ---

' Purpose: asserts that the condition is True
' 0.1.0    19.03.2026    Claude Code    Initially created
Public Sub AssertTrue _
( _
   ByVal b_arg_Condition As Boolean, _
   ByVal s_arg_Description As String _
)

   mAddResult b_arg_Condition, s_arg_Description, "AssertTrue", "True", CStr(b_arg_Condition)
End Sub

' Purpose: asserts that the condition is False
' 0.1.0    19.03.2026    Claude Code    Initially created
Public Sub AssertFalse _
( _
   ByVal b_arg_Condition As Boolean, _
   ByVal s_arg_Description As String _
)

   mAddResult Not b_arg_Condition, s_arg_Description, "AssertFalse", "False", CStr(b_arg_Condition)
End Sub

' Purpose: asserts that two values are equal (compared as strings for broad compatibility)
' 0.1.0    19.03.2026    Claude Code    Initially created
Public Sub AssertEqual _
( _
   ByVal v_arg_Expected As Variant, _
   ByVal v_arg_Actual As Variant, _
   ByVal s_arg_Description As String _
)

   Dim bPassed As Boolean
   Dim sExpected As String
   Dim sActual As String
   sExpected = mVariantToString(v_arg_Expected)
   sActual = mVariantToString(v_arg_Actual)
   bPassed = (sExpected = sActual)
   mAddResult bPassed, s_arg_Description, "AssertEqual", sExpected, sActual
End Sub

' Purpose: asserts that two values are not equal
' 0.1.0    19.03.2026    Claude Code    Initially created
Public Sub AssertNotEqual _
( _
   ByVal v_arg_NotExpected As Variant, _
   ByVal v_arg_Actual As Variant, _
   ByVal s_arg_Description As String _
)

   Dim bPassed As Boolean
   Dim sNotExpected As String
   Dim sActual As String
   sNotExpected = mVariantToString(v_arg_NotExpected)
   sActual = mVariantToString(v_arg_Actual)
   bPassed = (sNotExpected <> sActual)
   mAddResult bPassed, s_arg_Description, "AssertNotEqual", "Not " & sNotExpected, sActual
End Sub

' Purpose: asserts that the object reference is not Nothing
' 0.1.0    19.03.2026    Claude Code    Initially created
Public Sub AssertNotNothing _
( _
   ByVal o_arg_Object As Object, _
   ByVal s_arg_Description As String _
)

   Dim bPassed As Boolean
   bPassed = Not (o_arg_Object Is Nothing)
   Dim sActual As String
   If bPassed Then
      sActual = TypeName(o_arg_Object)
   Else
      sActual = "Nothing"
   End If
   mAddResult bPassed, s_arg_Description, "AssertNotNothing", "Not Nothing", sActual
End Sub

' Purpose: asserts that the object reference is Nothing
' 0.1.0    19.03.2026    Claude Code    Initially created
Public Sub AssertIsNothing _
( _
   ByVal o_arg_Object As Object, _
   ByVal s_arg_Description As String _
)

   Dim bPassed As Boolean
   bPassed = (o_arg_Object Is Nothing)
   Dim sActual As String
   If bPassed Then
      sActual = "Nothing"
   Else
      sActual = TypeName(o_arg_Object)
   End If
   mAddResult bPassed, s_arg_Description, "AssertIsNothing", "Nothing", sActual
End Sub

' Purpose: asserts that a Boolean-returning lower-level function succeeded (returned True)
' 0.1.0    19.03.2026    Claude Code    Initially created
Public Sub AssertSuccess _
( _
   ByVal b_arg_FunctionResult As Boolean, _
   ByVal s_arg_Description As String _
)

   mAddResult b_arg_FunctionResult, s_arg_Description, "AssertSuccess", "True (success)", CStr(b_arg_FunctionResult)
End Sub

' Purpose: asserts that a Boolean-returning lower-level function failed (returned False)
' 0.1.0    19.03.2026    Claude Code    Initially created
Public Sub AssertFailure _
( _
   ByVal b_arg_FunctionResult As Boolean, _
   ByVal s_arg_Description As String _
)

   mAddResult Not b_arg_FunctionResult, s_arg_Description, "AssertFailure", "False (failure)", CStr(b_arg_FunctionResult)
End Sub

' --- Private Helpers ---

' Purpose: creates and stores a single assertion result
' 0.1.0    19.03.2026    Claude Code    Initially created
Private Sub mAddResult _
( _
   ByVal b_arg_Passed As Boolean, _
   ByVal s_arg_Description As String, _
   ByVal s_arg_AssertType As String, _
   ByVal s_arg_Expected As String, _
   ByVal s_arg_Actual As String _
)

   Dim oC_Result As New DEV_f_C_AssertResult
   oC_Result.Initialize b_arg_Passed, s_arg_Description, s_arg_AssertType, s_arg_Expected, s_arg_Actual
   oCol_m_Results.Add oC_Result
   If b_arg_Passed Then
      l_m_PassedCount = l_m_PassedCount + 1
   Else
      l_m_FailedCount = l_m_FailedCount + 1
   End If
End Sub

' Purpose: converts a Variant to a string representation for comparison and display
' 0.1.0    19.03.2026    Claude Code    Initially created
Private Function mVariantToString(ByVal v_arg_Value As Variant) As String
   On Error Resume Next
   If IsNull(v_arg_Value) Then
      mVariantToString = "<Null>"
   ElseIf IsEmpty(v_arg_Value) Then
      mVariantToString = "<Empty>"
   ElseIf IsObject(v_arg_Value) Then
      mVariantToString = "<Object:" & TypeName(v_arg_Value) & ">"
   Else
      mVariantToString = CStr(v_arg_Value)
   End If
   If Err.Number <> 0 Then
      mVariantToString = "<Non-printable:" & TypeName(v_arg_Value) & ">"
      Err.Clear
   End If
End Function
