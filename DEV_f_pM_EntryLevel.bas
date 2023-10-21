Attribute VB_Name = "DEV_f_pM_EntryLevel"
' CORE-DEV - do not change, optionally remove when deploying app
'============================================================================================
'   NAME:     DEV_f_pM_EntryLevel
'============================================================================================
'   Purpose:  entry level procedures related to development
'   Access:   Private
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
'--------------------------------------------------------------------------------------------
'   BACKLOG
'   ''''''''''''''''''''
'   none
'============================================================================================
Option Explicit
Option Private Module

Private Const s_m_COMPONENT_NAME As String = "DEV_f_pM_EntryLevel"

' Purpose: exports all files required for version control to the project folder
' 0.14.0    31.01.2023    gueleh    Changed scope to public to call it from UI module, added export of wks names
'        and code lib reference data
' 0.12.0    16.08.2022    gueleh    Initially created
Public Sub DEV_f_p_ExportDataForVersionControl()

'Fixed, don't change
   Dim oC_Me As New f_C_CallParams
   oC_Me.s_prop_rw_ComponentName = s_m_COMPONENT_NAME
   
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Your custom settings here
   'Consult the manual for learning about the options for f_g_StartProcessing
   f_p_StartProcessing 'calling without args only inits the globals
   With oC_Me
      .s_prop_rw_ProcedureName = "DEV_f_p_ExportDataForVersionControl" 'Name of the sub
      .b_prop_rw_SilentError = False 'False will display a message box - you should only do this on entry level
      .s_prop_rw_ErrorMessage = "Exporting data for version control failed." 'A message that properly informs the user and the devs (silent errors will be logged nonetheless)
      .SetCallArgs "No args" 'If the sub takes args put the here like ("sExample:=" & sExample, "lExample:=" & lExample)
   End With
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

'Fixed, don't change
   If oC_f_p_FrameworkSettings.b_prop_rw_ThisIsATestRun Then f_p_RegisterUnitTest oC_Me
Try:
   On Error GoTo Catch
   
'>>>>>>> Your code here
      
   Dim oC_VersionControlExport As New DEV_f_C_VersionControlExport
   
      If Not _
   oC_VersionControlExport.bExportAllComponents() _
      Then Err.Raise _
         e_f_p_HandledError_ExecutionOfLowerLevelFunction, , _
         s_f_p_HandledErrorDescription(e_f_p_HandledError_ExecutionOfLowerLevelFunction)
   
      If Not _
   oC_VersionControlExport.bExportNameData() _
      Then Err.Raise _
         e_f_p_HandledError_ExecutionOfLowerLevelFunction, , _
         s_f_p_HandledErrorDescription(e_f_p_HandledError_ExecutionOfLowerLevelFunction)
   
      If Not _
   oC_VersionControlExport.bExportWorksheetNameData() _
      Then Err.Raise _
         e_f_p_HandledError_ExecutionOfLowerLevelFunction, , _
         s_f_p_HandledErrorDescription(e_f_p_HandledError_ExecutionOfLowerLevelFunction)
   
      If Not _
   oC_VersionControlExport.bExportReferenceData() _
      Then Err.Raise _
         e_f_p_HandledError_ExecutionOfLowerLevelFunction, , _
         s_f_p_HandledErrorDescription(e_f_p_HandledError_ExecutionOfLowerLevelFunction)

      If Not _
   oC_VersionControlExport.bExportSettingsSheetData() _
      Then Err.Raise _
         e_f_p_HandledError_ExecutionOfLowerLevelFunction, , _
         s_f_p_HandledErrorDescription(e_f_p_HandledError_ExecutionOfLowerLevelFunction)

      If Not _
   oC_VersionControlExport.bExportRangeContentData() _
      Then Err.Raise _
         e_f_p_HandledError_ExecutionOfLowerLevelFunction, , _
         s_f_p_HandledErrorDescription(e_f_p_HandledError_ExecutionOfLowerLevelFunction)


'End of your code <<<<<<<
   
'Fixed, don't change
Finally:
   On Error Resume Next

'>>>>>>> Your code here
   'everything that must be executed regardless of an error or not
'End of your code <<<<<<<

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Your custom settings here
   'Consult the manual for learning about the options for f_g_EndProcessing
   f_p_EndProcessing 'calling without args does nothing
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

'Fixed, don't change
   Exit Sub
Catch:
   If oC_Me.oC_prop_r_Error Is Nothing _
   Then f_p_RegisterError oC_Me, Err.Number, Err.Description
   If oC_f_p_FrameworkSettings.b_prop_rw_ThisIsATestRun Then f_p_RegisterExecutionError oC_Me
   If oC_f_p_FrameworkSettings.b_prop_r_DebugModeIsOn And Not oC_Me.b_prop_rw_ResumedOnce Then
      oC_Me.b_prop_rw_ResumedOnce = True: Stop: Resume
   Else
      f_p_HandleError oC_Me: Resume Finally
   End If
End Sub


