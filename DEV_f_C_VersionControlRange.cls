VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DEV_f_C_VersionControlRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private s_m_Name As String
Private s_m_NameOfRange As String
Private oRng_m_TheRange As Range

Public Sub Construct _
( _
   ByVal sName As String, _
   ByVal sNameOfRange As String, _
   ByRef oRng As Range _
)
   s_m_Name = sName
   s_m_NameOfRange = sNameOfRange
   oRng_m_TheRange = oRng
End Sub

Public Property Get s_prop_r_Name() As String
   s_prop_r_Name = s_m_Name
End Property

Public Property Get s_prop_r_NameOfRange() As String
   s_prop_r_NameOfRange = s_m_NameOfRange
End Property

Public Property Get oRng_prop_r_TheRange() As Range
   Set oRng_prop_r_TheRange = oRng_m_TheRange
End Property

