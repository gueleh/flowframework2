Attribute VB_Name = "DEV_f_pM_EntryLevel"
' -------------------------------------------------------------------------------------------
' CORE-DEV, do not change
'============================================================================================
'   NAME:     devfpMEntryLevel
'============================================================================================
'   Purpose:  entry level procedures related to development
'   Access:   Public
'   Type:     Module
'   Author:   Günther Lehner
'   Contact:  guleh@pm.me
'   GitHubID: gueleh
'   Required:
'   Usage:
'--------------------------------------------------------------------------------------------
'   VERSION HISTORY
'   Version    Date    Developer    Changes
'   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   0.12.0    16.08.2022    gueleh    Initially created
'--------------------------------------------------------------------------------------------
'   BACKLOG
'   ''''''''''''''''''''
'   none
'============================================================================================
Option Explicit
Option Private Module

Private Const s_m_COMPONENT_NAME As String = "devfpMEntryLevel"

' Purpose: exports all files required for version control to the project folder
' 0.14.0    31.01.2023    gueleh    Changed scope to public to call it from UI module, added export of wks names
'        and code lib reference data
' 0.12.0    16.08.2022    gueleh    Initially created
Public Sub DEV_f_p_ExportDataForVersionControl()

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'Change name of the sub if you want to have this information in the error log
   Const sNAME_OF_SUB As String = "f_p_TemplateSubEntryLevel"
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

'Fixed, don't change
   Dim bResumedOnce As Boolean
   Dim oCError As f_C_Error
   Dim bIsSilentError As Boolean
   Dim sErrorMessage As String
   
Try:
   On Error GoTo Catch
   af_p_StartProcessing 'turns off screen updating, unprotects sheets etc.
'End Fixed
   
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>> Your code here
   
   'Change this is necessary
   bIsSilentError = False 'if False, a MsgBox will be displayed. Otherwise only an entry in the error log.
   sErrorMessage = "The error message the user should see."
      
   Dim oC_VersionControlExport As New DEV_f_CVersionControlExport
   
      If Not _
   oC_VersionControlExport.bExportAllComponents() _
      Then Err.Raise l_f_p_EXECUTION_ERROR_NUMBER, , s_f_p_EXECUTION_ERROR_TEXT
   
      If Not _
   oC_VersionControlExport.bExportNameData() _
      Then Err.Raise l_f_p_EXECUTION_ERROR_NUMBER, , s_f_p_EXECUTION_ERROR_TEXT
   
      If Not _
   oC_VersionControlExport.bExportWorksheetNameData() _
      Then Err.Raise l_f_p_EXECUTION_ERROR_NUMBER, , s_f_p_EXECUTION_ERROR_TEXT
   
      If Not _
   oC_VersionControlExport.bExportReferenceData() _
      Then Err.Raise l_f_p_EXECUTION_ERROR_NUMBER, , s_f_p_EXECUTION_ERROR_TEXT

'End of your code <<<<<<<
   
'Fixed, don't change
Finally:
   On Error Resume Next
'End Fixed

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>> Your code here
'>>>>>>> everything that MUST BE EXECUTED regardless of an error or not


'End of your code <<<<<<<
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

'Fixed, don't change
   af_p_EndProcessing 'turns on screen updating, protects sheets etc.
   Exit Sub

Catch:
   'Set error data for logging if not already existent
   If oCError Is Nothing Then
      Set oCError = New f_C_Error
      oCError.SetErrorData _
         Err.Number, Err.Description, sNAME_OF_SUB, _
         s_m_COMPONENT_NAME, bIsSilentError, sErrorMessage
   End If
   
   'If in debug mode, then the code will stop once so that you can step into
   '  the row which caused the error.
   '  If already stopped and resumed once, the error data are sent to the error handler
   If b_p_DEBUG_MODE_IS_ON _
   And Not bResumedOnce Then
      bResumedOnce = True: Stop: Resume
   Else
      f_p_LogError oCError: Resume Finally
   End If
'End Fixed
End Sub


