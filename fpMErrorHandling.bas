Attribute VB_Name = "fpMErrorHandling"
' -------------------------------------------------------------------------------------------
' CORE, do not change
'============================================================================================
'   NAME:     fpMErrorHandling
'============================================================================================
'   Purpose:  error handling
'   Access:   Private
'   Type:     Module
'   Author:   G�nther Lehner
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

Private Const s_m_COMPONENT_NAME As String = "fpMErrorHandling"

'the error cases of the framework
Public Enum e_f_p_HandledErrors
   e_f_p_HandledError_GeneralError = 9999
   e_f_p_HandledError_AppSpecificError
   e_f_p_HandledError_ExecutionOfLowerLevelFunction
End Enum

' Purpose: returns the error description for the provided error cases
' 0.1.0    20220709    gueleh    Initially created
Public Function s_f_p_HandledErrorDescription( _
   ByVal efHandledError As e_f_p_HandledErrors, _
   Optional ByVal eafHandledError As e_af_p_HandledErrors = 19999) As String
   
   Dim sDesc As String
   Select Case efHandledError
      Case e_f_p_HandledError_AppSpecificError
         sDesc = s_af_p_HandledErrorDescription(eafHandledError)
      Case e_f_p_HandledError_GeneralError
         sDesc = "An error occurred. No specific description provided"
      Case e_f_p_HandledError_ExecutionOfLowerLevelFunction
         sDesc = "The execution of a lower level function failed, refer to error log."
      Case Else
         sDesc = "No description defined for this error. You can do this in Function s_af_p_HandledErrorDescription in module afpMErrorHandling."
   End Select
   s_f_p_HandledErrorDescription = sDesc
End Function

' Purpose: registers an error in the error stack
' 0.1.0    20220709    gueleh    Initially created
Public Sub f_p_RegisterError( _
   ByRef oC_f_Params As fCCallParams, _
   ByVal lErrorNumber As Long, _
   ByVal sErrorDescription As String)

   Dim oCError As New fCError
   With oCError
      .lErrorNumber = lErrorNumber
      .sErrorDescription = sErrorDescription
   End With
   oC_f_Params.SetoCError oCError
   If oCol_f_p_Errors.Count > 0 Then
      oCol_f_p_Errors.Add oC_f_Params, , 1
   Else
      oCol_f_p_Errors.Add oC_f_Params
   End If
   
End Sub

' Purpose: handles the error
' 0.1.0    20220709    gueleh    Initially created
Public Sub f_p_HandleError(ByRef oC_f_Params As fCCallParams)
   mLogError oC_f_Params
   With oC_f_Params
      If Not .bSilentError Then
         MsgBox .sErrorMessage, vbCritical
      End If
   End With
End Sub

' Purpose: adds entry to error log worksheet
' NOTE: the structure of the error log worksheet is hardcoded in this sub, any changes to it
'        do have to be reflected here accordingly
' 0.1.0    20220709    gueleh    Initially created
Private Sub mLogError(ByRef oC_f_Params As fCCallParams)
   Dim lRow As Long
   Const sANCHOR_ADDRESS As String = "A2"
   With afwksErrorLog
      lRow = .Range(sANCHOR_ADDRESS).CurrentRegion.Rows.Count + .Range(sANCHOR_ADDRESS).Row - 1
      .Cells(lRow, 1).Value2 = Format(Now(), "YYMMDD hh:mm:ss")
      .Cells(lRow, 2).Value2 = Environ("Username")
      .Cells(lRow, 3).Value2 = oC_f_Params.sComponentName
      .Cells(lRow, 4).Value2 = oC_f_Params.sProcedureName
      .Cells(lRow, 5).Value2 = oC_f_Params.oCError.lErrorNumber
      .Cells(lRow, 6).Value2 = oC_f_Params.oCError.sErrorDescription
      .Cells(lRow, 7).Value2 = oC_f_Params.bSilentError
      .Cells(lRow, 8).Value2 = oC_f_Params.sErrorMessage
      .Cells(lRow, 9).Value2 = oC_f_Params.sArgsAsString()
      .Cells.Calculate
   End With
End Sub
