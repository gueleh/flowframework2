Attribute VB_Name = "fpMErrorHandling"
' -------------------------------------------------------------------------------------------
' CORE, do not change
'============================================================================================
'   NAME:     fpMErrorHandling
'============================================================================================
'   Purpose:  error handling
'   Access:   Private
'   Type:     Module
'   Author:   Günther Lehner
'   Contact:  guenther.lehner@protonmail.com
'   GitHubID: gueleh
'   Required:
'   Usage:
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
Option Private Module

Private Const smCOMPONENT_NAME As String = "fpMErrorHandling"

'the error cases of the framework
Public Enum efHandledErrors
   efHandledErrorGeneralError = 9999
   efHandledErrorAppSpecificError
   efHandledErrorExecutionOfLowerLevelFunction
End Enum

' Purpose: returns the error description for the provided error cases
' 0.1.0    20220709    gueleh    Initially created
Public Function sfHandledErrorDescription( _
   ByVal efHandledError As efHandledErrors, _
   Optional ByVal eafHandledError As eafHandledErrors = 19999) As String
   
   Dim sDesc As String
   Select Case efHandledError
      Case efHandledErrorAppSpecificError
         sDesc = safHandledErrorDescription(eafHandledError)
      Case efHandledErrorGeneralError
         sDesc = "An error occurred. No specific description provided"
      Case efHandledErrorExecutionOfLowerLevelFunction
         sDesc = "The execution of a lower level function failed, refer to error log."
      Case Else
         sDesc = "No description defined for this error. You can do this in Function safHandledErrorDescription in module afpMErrorHandling."
   End Select
   sfHandledErrorDescription = sDesc
End Function

' Purpose: registers an error in the error stack
' 0.1.0    20220709    gueleh    Initially created
Public Sub fRegisterError( _
   ByRef oCParams As fCCallParams, _
   ByVal lErrorNumber As Long, _
   ByVal sErrorDescription As String)

   Dim oCError As New fCError
   With oCError
      .lErrorNumber = lErrorNumber
      .sErrorDescription = sErrorDescription
   End With
   oCParams.SetError oCError
   If colfgErrors.Count > 0 Then
      colfgErrors.Add oCParams, , 1
   Else
      colfgErrors.Add oCParams
   End If
   
End Sub

' Purpose: handles the error
' 0.1.0    20220709    gueleh    Initially created
Public Sub fHandleError(ByRef oCParams As fCCallParams)
   mLogError oCParams
   With oCParams
      If Not .bSilentError Then
         MsgBox .sErrorMessage, vbCritical
      End If
   End With
End Sub

' Purpose: adds entry to error log
' 0.1.0    20220709    gueleh    Initially created
Private Sub mLogError(ByRef oCParams As fCCallParams)
   Dim lRow As Long
   Const sANCHOR_ADDRESS As String = "A2"
   With afwksErrorLog
      lRow = .Range(sANCHOR_ADDRESS).CurrentRegion.Rows.Count + .Range(sANCHOR_ADDRESS).Row - 1
      .Cells(lRow, 1).Value2 = Format(Now(), "YYMMDD hh:mm:ss")
      .Cells(lRow, 2).Value2 = Environ("Username")
      .Cells(lRow, 3).Value2 = oCParams.sComponentName
      .Cells(lRow, 4).Value2 = oCParams.sProcedureName
      .Cells(lRow, 5).Value2 = oCParams.oCError.lErrorNumber
      .Cells(lRow, 6).Value2 = oCParams.oCError.sErrorDescription
      .Cells(lRow, 7).Value2 = oCParams.bSilentError
      .Cells(lRow, 8).Value2 = oCParams.sErrorMessage
      .Cells(lRow, 9).Value2 = oCParams.sArgsAsString()
      .Cells.Calculate
   End With
End Sub
