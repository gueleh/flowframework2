Attribute VB_Name = "DEV_f_pM_Test_UtilitiesRanges"
' -------------------------------------------------------------------------------------------
' DEV, remove from production version
'============================================================================================
'   NAME:     DEV_f_pM_Test_UtilitiesRanges
'============================================================================================
'   Purpose:  tests for f_pM_UtilitiesRanges
'   Access:   Private
'   Type:     Modul
'   Author:   WTS84036
'   Contact:  guleh@pm.me
'   GitHubID: gueleh
'   Required:
'   Usage:
'--------------------------------------------------------------------------------------------
'   VERSION HISTORY
'   Version    Date    Developer    Changes
'   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   1.16    16.08.2024    gueleh    Initially created
'--------------------------------------------------------------------------------------------
'   BACKLOG
'   ''''''''''''''''''''
'   none
'============================================================================================
Option Explicit
Option Private Module

Private Const s_m_COMPONENT_NAME As String = "DEV_f_pM_Test_UtilitiesRanges"

'Purpose: simple test based on Debug.Assert, i.e. code execution will stop if assertions fails
Private Sub mTest_b_f_p_RangeSizeAndCellContentsAreEqual()
'Fixed, don't change
   Dim oC_Me As New f_C_CallParams: oC_Me.s_prop_rw_ComponentName = s_m_COMPONENT_NAME
'>>>>>>> Your custom settings here
   f_p_StartProcessing e_f_p_ProcessingMode_AutoCalcOffOnSceenUpdatingOffOn
   With oC_Me
      .s_prop_rw_ProcedureName = "f_p_TemplateSubEntryLevelCompact" 'Name of the sub
      .b_prop_rw_SilentError = False 'False will display a message box - you should only do this on entry level
      .s_prop_rw_ErrorMessage = "Your message here." 'A message that properly informs the user and the devs (silent errors will be logged nonetheless)
      .SetCallArgs "No args" 'If the sub takes args put the here like ("sExample:=" & sExample, "lExample:=" & lExample)
   End With
'Fixed, don't change
   If oC_f_p_FrameworkSettings.b_prop_rw_ThisIsATestRun Then f_p_RegisterUnitTest oC_Me
Try: On Error GoTo Catch


'>>>>>>> Your code here
   Dim oRngOne As Range
   Dim oRngTwo As Range
   Dim bResult As Boolean
   Dim sResult As String
   
   'Processing error
   Debug.Assert Not b_f_p_RangeSizeAndCellContentsAreEqual(bResult, oRngOne, oRngTwo, sResult)
   
   'Processing error
   DEV_Reset_DEV_f_wks_TestCanvas
   Set oRngOne = DEV_f_wks_TestCanvas.Range("A1:E1")
   Set oRngTwo = DEV_f_wks_TestCanvas.Range("A2:D2")
   Debug.Assert b_f_p_RangeSizeAndCellContentsAreEqual(bResult, oRngOne, oRngTwo, sResult)
   Debug.Assert Not bResult
   
   'Same size and empty
   Set oRngTwo = DEV_f_wks_TestCanvas.Range("A2:E2")
   Debug.Assert b_f_p_RangeSizeAndCellContentsAreEqual(bResult, oRngOne, oRngTwo, sResult)
   Debug.Assert bResult
   
   'Same size different cell contents
   DEV_f_wks_TestCanvas.Range("B1").Value2 = "Test"
   Debug.Assert b_f_p_RangeSizeAndCellContentsAreEqual(bResult, oRngOne, oRngTwo, sResult)
   Debug.Assert Not bResult

   'Same size, same contents
   DEV_f_wks_TestCanvas.Range("B2").Value2 = "Test"
   Debug.Assert b_f_p_RangeSizeAndCellContentsAreEqual(bResult, oRngOne, oRngTwo, sResult)
   Debug.Assert bResult


'End of your code <<<<<<<
   

'Fixed, don't change
Finally: On Error Resume Next


'>>>>>>> Your code here



'End of your code <<<<<<<


'>>>>>>> Your custom settings here
   f_p_EndProcessing e_f_p_ProcessingMode_AutoCalcOffOnSceenUpdatingOffOn
'Fixed, don't change
   Exit Sub
HandleError: af_pM_ErrorHandling.af_p_Hook_ErrorHandling_EntryLevel


'>>>>>>> Your code here



'End of your code <<<<<<<
   
   
'Fixed, don't change
   Resume Finally
Catch:
   If oC_Me.oC_prop_r_Error Is Nothing Then f_p_RegisterError oC_Me, Err.Number, Err.Description
   If oC_f_p_FrameworkSettings.b_prop_rw_ThisIsATestRun Then f_p_RegisterExecutionError oC_Me
   If oC_f_p_FrameworkSettings.b_prop_r_DebugModeIsOn And Not oC_Me.b_prop_rw_ResumedOnce Then
      oC_Me.b_prop_rw_ResumedOnce = True: Stop: Resume
   Else
      f_p_HandleError oC_Me: GoTo HandleError
   End If

End Sub
