Attribute VB_Name = "f_pM_ErrorHandling"
' CORE, do not change
'============================================================================================
'   NAME:     f_pM_ErrorHandling
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
' 0.10.0    04.08.2022    gueleh    Changed property names to meet new convention
'   0.1.0    20220709    gueleh    Initially created
'--------------------------------------------------------------------------------------------
'   BACKLOG
'   ''''''''''''''''''''
'   none
'============================================================================================
Option Explicit
Option Private Module

Private Const s_m_COMPONENT_NAME As String = "f_pM_ErrorHandling"

'the error cases of the framework
Public Enum e_f_p_HandledErrors
   e_f_p_HandledError_GeneralError = 9999
   e_f_p_HandledError_AppSpecificError
   e_f_p_HandledError_ExecutionOfLowerLevelFunction
End Enum

' Purpose: returns the error description for the provided error cases
' 0.1.0    20220709    gueleh    Initially created
Public Function s_f_p_HandledErrorDescription _
( _
   ByVal ef_arg_HandledError As e_f_p_HandledErrors, _
   Optional ByVal eaf_arg_HandledError As e_af_p_HandledErrors = 19999 _
) As String
   
   Dim sDesc As String
   Select Case ef_arg_HandledError
      Case e_f_p_HandledError_AppSpecificError
         sDesc = s_af_p_HandledErrorDescription(eaf_arg_HandledError)
      Case e_f_p_HandledError_GeneralError
         sDesc = "An error occurred. No specific description provided"
      Case e_f_p_HandledError_ExecutionOfLowerLevelFunction
         sDesc = "The execution of a lower level function failed, refer to error log."
      Case Else
         sDesc = "No description defined for this error. You can do this in Function s_af_p_HandledErrorDescription in module a_f_pM_ErrorHandling."
   End Select
   s_f_p_HandledErrorDescription = sDesc
End Function

' Purpose: registers an error in the error stack
' 0.1.0    20220709    gueleh    Initially created
Public Sub f_p_RegisterError _
( _
   ByRef oC_f_arg_Params As f_C_CallParams, _
   ByVal l_arg_ErrorNumber As Long, _
   ByVal s_arg_ErrorDescription As String _
)

   Dim oCError As New f_C_Error
   With oCError
      .l_prop_rw_ErrorNumber = l_arg_ErrorNumber
      .s_prop_rw_ErrorDescription = s_arg_ErrorDescription
   End With
   oC_f_arg_Params.SetoCError oCError
   If oCol_f_p_Errors.Count > 0 Then
      oCol_f_p_Errors.Add oC_f_arg_Params, , 1
   Else
      oCol_f_p_Errors.Add oC_f_arg_Params
   End If
   
End Sub

' Purpose: handles the error
' 0.1.0    20220709    gueleh    Initially created
Public Sub f_p_HandleError _
( _
   ByRef oC_f_arg_Params As f_C_CallParams _
)
   
   mLogError oC_f_arg_Params
   With oC_f_arg_Params
      If Not .b_prop_rw_SilentError Then
         MsgBox .s_prop_rw_ErrorMessage, vbCritical
      End If
   End With
End Sub

' Purpose: adds entry to error log worksheet
' NOTE: the structure of the error log worksheet is hardcoded in this sub, any changes to it
'        do have to be reflected here accordingly
' 0.1.0    20220709    gueleh    Initially created
Private Sub mLogError _
( _
   ByRef oC_f_arg_Params As f_C_CallParams _
)
   
   Dim lRow As Long
   Const sANCHOR_ADDRESS As String = "A2"
   With af_wks_ErrorLog
      lRow = .Range(sANCHOR_ADDRESS).CurrentRegion.Rows.Count + .Range(sANCHOR_ADDRESS).Row - 1
      .Cells(lRow, 1).Value2 = Format(Now(), "YYMMDD hh:mm:ss")
      .Cells(lRow, 2).Value2 = Environ("Username")
      .Cells(lRow, 3).Value2 = oC_f_arg_Params.s_prop_rw_ComponentName
      .Cells(lRow, 4).Value2 = oC_f_arg_Params.s_prop_rw_ProcedureName
      .Cells(lRow, 5).Value2 = oC_f_arg_Params.oC_prop_r_Error.l_prop_rw_ErrorNumber
      .Cells(lRow, 6).Value2 = oC_f_arg_Params.oC_prop_r_Error.s_prop_rw_ErrorDescription
      .Cells(lRow, 7).Value2 = oC_f_arg_Params.b_prop_rw_SilentError
      .Cells(lRow, 8).Value2 = oC_f_arg_Params.s_prop_rw_ErrorMessage
      .Cells(lRow, 9).Value2 = oC_f_arg_Params.sArgsAsString()
      .Cells.Calculate
   End With
End Sub
