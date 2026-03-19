VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DEV_f_C_AssertResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' CORE-DEV - do not change, optionally remove when deploying app
'============================================================================================
'   NAME:     DEV_f_C_AssertResult
'============================================================================================
'   Purpose:  Stores the result of a single assertion within a unit test
'   Access:   Public
'   Type:     Class Module
'   Author:   Claude Code
'   Contact:  guenther.lehner@protonmail.com
'   GitHubID: gueleh
'   Required:
'   Usage:    Created internally by DEV_f_C_Assert for each assertion call
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

Private Const s_m_COMPONENT_NAME As String = "DEV_f_C_AssertResult"

Private b_m_Passed As Boolean
Private s_m_Description As String
Private s_m_Expected As String
Private s_m_Actual As String
Private s_m_AssertType As String

Public Property Get b_prop_r_Passed() As Boolean: b_prop_r_Passed = b_m_Passed: End Property

Public Property Get s_prop_r_Description() As String: s_prop_r_Description = s_m_Description: End Property

Public Property Get s_prop_r_Expected() As String: s_prop_r_Expected = s_m_Expected: End Property

Public Property Get s_prop_r_Actual() As String: s_prop_r_Actual = s_m_Actual: End Property

Public Property Get s_prop_r_AssertType() As String: s_prop_r_AssertType = s_m_AssertType: End Property

' Purpose: initializes the assert result with all relevant data
' 0.1.0    19.03.2026    Claude Code    Initially created
Public Sub Initialize _
( _
   ByVal b_arg_Passed As Boolean, _
   ByVal s_arg_Description As String, _
   ByVal s_arg_AssertType As String, _
   Optional ByVal s_arg_Expected As String = vbNullString, _
   Optional ByVal s_arg_Actual As String = vbNullString _
)

   b_m_Passed = b_arg_Passed
   s_m_Description = s_arg_Description
   s_m_AssertType = s_arg_AssertType
   s_m_Expected = s_arg_Expected
   s_m_Actual = s_arg_Actual
End Sub

' Purpose: returns a formatted single-line summary of the assertion result
' 0.1.0    19.03.2026    Claude Code    Initially created
Public Function s_prop_r_Summary() As String
   Dim sResult As String
   If b_m_Passed Then
      sResult = "[PASS] "
   Else
      sResult = "[FAIL] "
   End If
   sResult = sResult & s_m_AssertType & ": " & s_m_Description
   If Not b_m_Passed And Len(s_m_Expected) > 0 Then
      sResult = sResult & " (Expected: " & s_m_Expected & " | Actual: " & s_m_Actual & ")"
   End If
   s_prop_r_Summary = sResult
End Function
