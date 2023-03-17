Attribute VB_Name = "f_pM_ErrorHandling"
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
' 0.1.0    17.03.2023    gueleh    Initially created
'--------------------------------------------------------------------------------------------
'   BACKLOG
'   ''''''''''''''''''''
'   none
'============================================================================================
Option Explicit
Option Private Module

Private Const s_m_COMPONENT_NAME As String = "fpMErrorHandling"

' Purpose: adds entry to error log worksheet
' NOTE: the structure of the error log worksheet is hardcoded in this sub, any changes to it
'        do have to be reflected here accordingly
' 0.1.0    17.03.2023    gueleh    Initially created
Public Sub f_p_LogError(ByRef oCError As f_C_Error)
   Dim lRow As Long
   Const sANCHOR_ADDRESS As String = "A2"
   With af_wks_ErrorLog
      lRow = .Range(sANCHOR_ADDRESS).CurrentRegion.Rows.Count + .Range(sANCHOR_ADDRESS).Row - 1
      .Cells(lRow, 1).Value2 = Format(Now(), "YYMMDD hh:mm:ss")
      .Cells(lRow, 2).Value2 = Environ("Username")
      .Cells(lRow, 3).Value2 = oCError.s_prop_rw_NameOfModule
      .Cells(lRow, 4).Value2 = oCError.s_prop_rw_NameOfModule
      .Cells(lRow, 5).Value2 = oCError.l_prop_rw_ErrorNumber
      .Cells(lRow, 6).Value2 = oCError.s_prop_rw_ErrorDescription
      .Cells(lRow, 7).Value2 = oCError.b_prop_rw_IsSilentError
      .Cells(lRow, 8).Value2 = oCError.s_prop_rw_ErrorMessage
      .Cells(lRow).Calculate
   End With
   If Not oCError.b_prop_rw_IsSilentError Then
      MsgBox oCError.s_prop_rw_ErrorMessage, vbCritical, ThisWorkbook.Name
   End If
End Sub
