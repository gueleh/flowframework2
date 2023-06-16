VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "f_C_Error"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Purpose: CORE, do not change - trivial class to store error object properties
' 0.10.0    04.08.2022    gueleh    Changed property names to meet new convention
' 0.7.0    02.08.2022    gueleh    Refactored to use properties instead of direct access to fields
' 0.1.0    20220709    gueleh    Initially created
Option Explicit

Private l_m_ErrorNumber As Long
Private s_m_ErrorDescription As String

Public Property Get l_prop_rw_ErrorNumber() As Long: l_prop_rw_ErrorNumber = l_m_ErrorNumber: End Property

Public Property Let l_prop_rw_ErrorNumber(ByVal lNewValue As Long): l_m_ErrorNumber = lNewValue: End Property

Public Property Get s_prop_rw_ErrorDescription() As String: s_prop_rw_ErrorDescription = s_m_ErrorDescription: End Property

Public Property Let s_prop_rw_ErrorDescription(ByVal sNewValue As String): s_m_ErrorDescription = sNewValue: End Property
