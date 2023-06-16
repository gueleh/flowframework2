VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "f_C_CallParams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' CORE, do not change
'============================================================================================
'   NAME:     f_C_CallParams
'============================================================================================
'   Purpose:  instantiated in each called non-trivial procedure, containing meta data
'           for processing, testing, error handling
'   Access:   Public
'   Type:     Class Module
'   Author:   Günther Lehner
'   Contact:  guenther.lehner@protonmail.com
'   GitHubID: gueleh
'   Required:
'   Usage: consult the template procedures in module fmTemplatesCore to learn about the usage
'--------------------------------------------------------------------------------------------
'   VERSION HISTORY
'   Version    Date    Developer    Changes
'   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 0.10.0    04.08.2022    gueleh    Changed property names to meet new convention
' 0.7.0    02.08.2022    gueleh    Just renaming
' 0.5.0    15.07.2022    gueleh    Changed public fields to private fields and added properties
'   0.1.0    20220709    gueleh    Initially created
'--------------------------------------------------------------------------------------------
'   BACKLOG
'   ''''''''''''''''''''
'   none
'============================================================================================
Option Explicit

Private Const s_m_COMPONENT_NAME As String = "f_C_CallParams"

Private s_m_ProcedureName As String ' the name of the procedure the class instance refers to
Private s_m_ComponentName As String ' the name of the code component the procedure belongs to
Private l_m_UnitTestIndex As Long ' the position of the test in the collection of unit tests, _
   for direct reference to it
Private b_m_ResumedOnce As Boolean ' used for stopping and resuming once in an erroneous procedure _
   when in debug mode
Private b_m_SilentError As Boolean ' if true, no error message will be displayed for an error occuring _
   in the procedure the instance of this class refers to
Private s_m_ErrorMessage As String
Private oC_m_Error As f_C_Error
Private va_m_Args() As Variant ' supposed to store the args passed in to the procedure, for test and error documentation

Public Property Get b_prop_rw_ResumedOnce() As Boolean: b_prop_rw_ResumedOnce = b_m_ResumedOnce: End Property

Public Property Let b_prop_rw_ResumedOnce(ByVal bNewValue As Boolean): b_m_ResumedOnce = bNewValue: End Property

Public Property Get b_prop_rw_SilentError() As Boolean: b_prop_rw_SilentError = b_m_SilentError: End Property

Public Property Let b_prop_rw_SilentError(ByVal bNewValue As Boolean): b_m_SilentError = bNewValue: End Property

Public Property Get l_prop_rw_UnitTestIndex() As Long: l_prop_rw_UnitTestIndex = l_m_UnitTestIndex: End Property

Public Property Let l_prop_rw_UnitTestIndex(ByVal lNewValue As Long): l_m_UnitTestIndex = lNewValue: End Property

Public Property Get oC_prop_r_Error() As f_C_Error: Set oC_prop_r_Error = oC_m_Error: End Property

Public Property Get s_prop_rw_ComponentName() As String: s_prop_rw_ComponentName = s_m_ComponentName: End Property

Public Property Let s_prop_rw_ComponentName(ByVal sNewValue As String): s_m_ComponentName = sNewValue: End Property

Public Property Get s_prop_rw_ErrorMessage() As String: s_prop_rw_ErrorMessage = s_m_ErrorMessage: End Property

Public Property Let s_prop_rw_ErrorMessage(ByVal sNewValue As String): s_m_ErrorMessage = sNewValue: End Property

Public Property Get s_prop_rw_ProcedureName() As String: s_prop_rw_ProcedureName = s_m_ProcedureName: End Property

Public Property Let s_prop_rw_ProcedureName(ByVal sNewValue As String): s_m_ProcedureName = sNewValue: End Property

Public Property Get va_prop_r_Args() As Variant(): va_prop_r_Args = va_m_Args: End Property

' Purpose: returns the args as list with comma separation, for non-printable args the type is provided
' 0.1.0    20220709    gueleh    Initially created
Public Function sArgsAsString() As String
   Dim l As Long
   Dim sResult As String
   Dim sElement As String
   If Not IsEmpty(va_m_Args) Then
      For l = LBound(va_m_Args) To UBound(va_m_Args)
         On Error Resume Next
         sElement = va_m_Args(l)
         If Err.Number > 0 Then
            sElement = "Not prinable, type: " & TypeName(va_m_Args(l))
            Err.Clear
         End If
         sResult = sElement & ", "
      Next l
      sResult = Left$(sResult, Len(sResult) - 2)
   End If
   sArgsAsString = sResult
End Function

' Purpose: stores the passed in argument information in an array for informational purposes,
'        e.g. the error log can contain the params passed to the erroneous procedure
' 0.5.0    15.07.2022    gueleh    Initially documented
Public Sub SetCallArgs(ParamArray vaArgs() As Variant)
   On Error Resume Next
   va_m_Args = vaArgs
End Sub

' Purpose: sets the oC_prop_r_Error instance of the class, sub instead of property for easier handling
' 0.5.0    15.07.2022    gueleh    Initially documented and renamed
Public Sub SetoCError(ByRef oCError As f_C_Error)
   Set oC_m_Error = oCError
End Sub
