VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "fCCallParams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' -------------------------------------------------------------------------------------------
' CORE, do not change
'============================================================================================
'   NAME:     fCCallParams
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
' 0.5.0    15.07.2022    gueleh    Changed public fields to private fields and added properties
'   0.1.0    20220709    gueleh    Initially created
'--------------------------------------------------------------------------------------------
'   BACKLOG
'   ''''''''''''''''''''
'   none
'============================================================================================
Option Explicit

Private Const smCOMPONENT_NAME As String = "fCCallParams"

Private smProcedureName As String ' the name of the procedure the class instance refers to
Private smComponentName As String ' the name of the code component the procedure belongs to
Private lmUnitTestIndex As Long ' the position of the test in the collection of unit tests, _
   for direct reference to it
Private bmResumedOnce As Boolean ' used for stopping and resuming once in an erroneous procedure _
   when in debug mode
Private bmSilentError As Boolean ' if true, no error message will be displayed for an error occuring _
   in the procedure the instance of this class refers to
Private smErrorMessage As String
Private omCError As fCError
Private vamArgs() As Variant ' supposed to store the args passed in to the procedure, for test and error documentation

Public Property Get bResumedOnce() As Boolean: bResumedOnce = bmResumedOnce: End Property

Public Property Let bResumedOnce(ByVal bNewValue As Boolean): bmResumedOnce = bNewValue: End Property

Public Property Get bSilentError() As Boolean: bSilentError = bmSilentError: End Property

Public Property Let bSilentError(ByVal bNewValue As Boolean): bmSilentError = bNewValue: End Property

Public Property Get lUnitTestIndex() As Long: lUnitTestIndex = lmUnitTestIndex: End Property

Public Property Let lUnitTestIndex(ByVal lNewValue As Long): lmUnitTestIndex = lNewValue: End Property

Public Property Get oCError() As fCError: Set oCError = omCError: End Property

Public Property Get sComponentName() As String: sComponentName = smComponentName: End Property

Public Property Let sComponentName(ByVal sNewValue As String): smComponentName = sNewValue: End Property

Public Property Get sErrorMessage() As String: sErrorMessage = smErrorMessage: End Property

Public Property Let sErrorMessage(ByVal sNewValue As String): smErrorMessage = sNewValue: End Property

Public Property Get sProcedureName() As String: sProcedureName = smProcedureName: End Property

Public Property Let sProcedureName(ByVal sNewValue As String): smProcedureName = sNewValue: End Property

Public Property Get vaArgs() As Variant(): vaArgs = vamArgs: End Property

' Purpose: returns the args as list with comma separation, for non-printable args the type is provided
' 0.1.0    20220709    gueleh    Initially created
Public Function sArgsAsString() As String
   Dim l As Long
   Dim sResult As String
   Dim sElement As String
   If Not IsEmpty(vamArgs) Then
      For l = LBound(vamArgs) To UBound(vamArgs)
         On Error Resume Next
         sElement = vamArgs(l)
         If Err.Number > 0 Then
            sElement = "Not prinable, type: " & TypeName(vamArgs(l))
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
   vamArgs = vaArgs
End Sub

' Purpose: sets the oCError instance of the class, sub instead of property for easier handling
' 0.5.0    15.07.2022    gueleh    Initially documented and renamed
Public Sub SetoCError(ByRef oCError As fCError)
   Set omCError = oCError
End Sub
