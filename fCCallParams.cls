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
'   Purpose:  instantiated in each called non-trivial procedure, for processing, testing, error handling
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
'   0.1.0    20220709    gueleh    Initially created
'--------------------------------------------------------------------------------------------
'   BACKLOG
'   ''''''''''''''''''''
'   none
'============================================================================================
Option Explicit

Private Const smCOMPONENT_NAME As String = "fCCallParams"

Public sComponentName As String
Public sProcedureName As String
Public lUnitTestIndex As Long ' the position of the test in the collection of unit tests, for direct reference to it
Public bResumedOnce As Boolean
Public bSilentError As Boolean
Public sErrorMessage As String
Private omCError As fCError
Private vamArgs() As Variant ' supposed to store the args passed in to the procedure, for test and error documentation

Public Property Get vaArgs() As Variant()
   vaArgs = vamArgs
End Property

Public Property Get oCError() As fCError
   Set oCError = omCError
End Property
 
Public Sub SetCallArgs(ParamArray vaArgs() As Variant)
   On Error Resume Next
   vamArgs = vaArgs
End Sub

Public Sub SetError(ByRef oCError As fCError)
   Set omCError = oCError
End Sub

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
