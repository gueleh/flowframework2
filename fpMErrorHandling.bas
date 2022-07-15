Attribute VB_Name = "fmErrorHandling"
' -------------------------------------------------------------------------------------------
' CORE, do not change
'============================================================================================
'   NAME:     fmErrorHandling
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

Private Const msCOMPONENT_NAME As String = "fmErrorHandling"

'the error cases of the framework
Public Enum fHandledErrors
   fHandledErrorGeneralError = 9999
   fHandledErrorAppSpecificError
   fHandledErrorExecutionOfLowerLevelFunction
End Enum

' Purpose: returns the error description for the provided error cases
' 0.1.0    20220709    gueleh    Initially created
Public Function fsHandledErrorDescription( _
   ByVal feHandledError As fHandledErrors, _
   Optional ByVal afeHandledError As afHandledErrors = 19999) As String
   
   Dim sDesc As String
   Select Case feHandledError
      Case fHandledErrorAppSpecificError
         sDesc = afsHandledErrorDescription(afeHandledError)
      Case fHandledErrorGeneralError
         sDesc = "An error occurred. No specific description provided"
      Case fHandledErrorExecutionOfLowerLevelFunction
         sDesc = "The execution of a lower level function failed, refer to error log."
      Case Else
         sDesc = "No description defined for this error. You can do this in Function afsHandledErrorDescription in module afmErrorHandling."
   End Select
   fsHandledErrorDescription = sDesc
End Function

' Purpose: registers an error in the error stack
' 0.1.0    20220709    gueleh    Initially created
Public Sub fRegisterError( _
   ByRef clsParams As fclsCallParams, _
   ByVal lErrorNumber As Long, _
   ByVal sErrorDescription As String)

   Dim clsError As New fclsError
   With clsError
      .lErrorNumber = lErrorNumber
      .sErrorDescription = sErrorDescription
   End With
   clsParams.SetError clsError
   If fgcolErrors.Count > 0 Then
      fgcolErrors.Add clsParams, , 1
   Else
      fgcolErrors.Add clsParams
   End If
   
End Sub

' Purpose: handles the error
' 0.1.0    20220709    gueleh    Initially created
Public Sub fHandleError(ByRef clsParams As fclsCallParams)
   mLogError clsParams
   With clsParams
      If Not .bSilentError Then
         MsgBox .sErrorMessage, vbCritical
      End If
   End With
End Sub

' Purpose: adds entry to error log
' 0.1.0    20220709    gueleh    Initially created
Private Sub mLogError(ByRef clsParams As fclsCallParams)
   Dim lRow As Long
   Const sANCHOR_ADDRESS As String = "A2"
   With afwksErrorLog
      lRow = .Range(sANCHOR_ADDRESS).CurrentRegion.Rows.Count + .Range(sANCHOR_ADDRESS).Row - 1
      .Cells(lRow, 1).Value2 = Format(Now(), "YYMMDD hh:mm:ss")
      .Cells(lRow, 2).Value2 = Environ("Username")
      .Cells(lRow, 3).Value2 = clsParams.sComponentName
      .Cells(lRow, 4).Value2 = clsParams.sProcedureName
      .Cells(lRow, 5).Value2 = clsParams.clsError.lErrorNumber
      .Cells(lRow, 6).Value2 = clsParams.clsError.sErrorDescription
      .Cells(lRow, 7).Value2 = clsParams.bSilentError
      .Cells(lRow, 8).Value2 = clsParams.sErrorMessage
      .Cells(lRow, 9).Value2 = clsParams.sArgsAsString()
      .Cells.Calculate
   End With
End Sub
