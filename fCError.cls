VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "fCError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Purpose: CORE, do not change - trivial class to store error object properties
' 0.7.0    02.08.2022    gueleh    Refactored to use properties instead of direct access to fields
' 0.1.0    20220709    gueleh    Initially created
Option Explicit

Private l_m_ErrorNumber As Long
Private s_m_ErrorDescription As String

Public Property Get lErrorNumber() As Long: lErrorNumber = l_m_ErrorNumber: End Property

Public Property Let lErrorNumber(ByVal lNewValue As Long): l_m_ErrorNumber = lNewValue: End Property

Public Property Get sErrorDescription() As String: sErrorDescription = s_m_ErrorDescription: End Property

Public Property Let sErrorDescription(ByVal sNewValue As String): s_m_ErrorDescription = sNewValue: End Property
