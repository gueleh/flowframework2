VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "f_C_Error"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Purpose: stores error data for error log
' 0.1.0    17.03.2023    gueleh    Initially created
Option Explicit

Private l_m_ErrorNumber As Long
Private s_m_ErrorDescription As String
Private s_m_NameOfMethod As String
Private s_m_NameOfModule As String
Private b_m_IsSilentError As Boolean
Private s_m_ErrorMessage As String

Public Property Get l_prop_rw_ErrorNumber() As Long
   l_prop_rw_ErrorNumber = l_m_ErrorNumber
End Property

Public Property Let l_prop_rw_ErrorNumber(ByVal lNewValue As Long)
   l_m_ErrorNumber = lNewValue
End Property

Public Property Get s_prop_rw_ErrorDescription() As String
   s_prop_rw_ErrorDescription = s_m_ErrorDescription
End Property

Public Property Let s_prop_rw_ErrorDescription(ByVal sNewValue As String)
   s_m_ErrorDescription = sNewValue
End Property

Public Property Get s_prop_rw_NameOfMethod() As String
   s_prop_rw_NameOfMethod = s_m_NameOfMethod
End Property

Public Property Let s_prop_rw_NameOfMethod(ByVal sNewValue As String)
   s_m_NameOfMethod = sNewValue
End Property

Public Property Get s_prop_rw_NameOfModule() As String
   s_prop_rw_NameOfModule = s_m_NameOfModule
End Property

Public Property Let s_prop_rw_NameOfModule(ByVal sNewValue As String)
   s_m_NameOfModule = sNewValue
End Property

Public Property Get b_prop_rw_IsSilentError() As Boolean
   b_prop_rw_IsSilentError = b_m_IsSilentError
End Property

Public Property Let b_prop_rw_IsSilentError(ByVal bNewValue As Boolean)
   b_m_IsSilentError = bNewValue
End Property

Public Property Get s_prop_rw_ErrorMessage() As String
   s_prop_rw_ErrorMessage = s_m_ErrorMessage
End Property

Public Property Let s_prop_rw_ErrorMessage(ByVal sNewValue As String)
   s_m_ErrorMessage = sNewValue
End Property

Private Sub Class_Initialize()
   b_m_IsSilentError = True
End Sub

Public Sub SetErrorData( _
   ByVal lErrorNumber As Long, _
   ByVal sErrorDescription As String, _
   ByVal sNameOfMethod As String, _
   ByVal sNameOfModule As String, _
   Optional ByVal bIsSilentError = True, _
   Optional ByVal sErrorMessage = "No error message defined.")

   l_m_ErrorNumber = lErrorNumber
   s_m_ErrorDescription = sErrorDescription
   s_m_NameOfMethod = sNameOfMethod
   s_m_NameOfModule = sNameOfModule
   b_m_IsSilentError = bIsSilentError
   s_m_ErrorMessage = sErrorMessage

End Sub
