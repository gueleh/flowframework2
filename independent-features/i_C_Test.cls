VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "i_C_Test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private s_m_ID As String
Private s_m_Description As String
Private s_m_Input As String
Private s_m_ExpectedOutput As String
Private s_m_StartingState As String
Private b_m_TestPassed As Boolean
Private s_m_ActualOutput As String

Public Sub AddTest(ByVal sID As String, ByVal sDescription As String, ByVal sINPUT As String, ByVal sExpectedOutput As String, ByVal sStartingState As String, Optional ByVal bPassed As Boolean, Optional ByVal sActualOutput As String = "Not provided")
   s_m_ID = sID
   s_m_Description = sDescription
   s_m_Input = sINPUT
   s_m_ExpectedOutput = sExpectedOutput
   s_m_StartingState = sStartingState
   b_m_TestPassed = bPassed
   s_m_ActualOutput = sActualOutput
End Sub

Public Sub SetResult(ByVal bPassed As Boolean)
   b_m_TestPassed = bPassed
End Sub

Public Sub SetActualOutput(ByVal sActualOutput As String)
   s_m_ActualOutput = sActualOutput
End Sub

Public Function vaGetTestData() As Variant
   Dim vaData(1 To 1, 1 To 7) As Variant
   vaData(1, 1) = s_m_ID
   vaData(1, 2) = s_m_Description
   vaData(1, 3) = s_m_Input
   vaData(1, 4) = s_m_ExpectedOutput
   vaData(1, 5) = s_m_ActualOutput
   vaData(1, 6) = s_m_StartingState
   vaData(1, 7) = b_m_TestPassed
   vaGetTestData = vaData
End Function

Public Function bPassed() As Boolean
   bPassed = b_m_TestPassed
End Function
