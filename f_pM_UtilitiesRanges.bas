Attribute VB_Name = "f_pM_UtilitiesRanges"
' CORE, do not change
'============================================================================================
'   NAME:     f_pM_UtilitiesRanges
'============================================================================================
'   Purpose:  utilities for working with ranges
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
'   1.16.0    16.08.2024    gueleh    Initially created
'--------------------------------------------------------------------------------------------
'   BACKLOG
'   ''''''''''''''''''''
'   none
'============================================================================================
Option Explicit
Option Private Module

Private Const s_m_COMPONENT_NAME As String = "f_pM_UtilitiesRanges"

' Purpose: template for a non-trivial lower level procedure with error handling and execution control, compact version
Public Function b_f_p_RangeSizeAndCellContentsAreEqual(ByRef bAreEqual As Boolean, ByRef oRngOne As Range, ByRef oRngTwo As Range, ByRef sStatusMessage As String) As Boolean
'Fixed, don't change
   Dim oC_Me As New f_C_CallParams: oC_Me.s_prop_rw_ComponentName = s_m_COMPONENT_NAME: If oC_f_p_FrameworkSettings.b_prop_rw_ThisIsATestRun Then f_p_RegisterUnitTest oC_Me
'>>>>>>> Your custom settings here
   With oC_Me
      .s_prop_rw_ProcedureName = "b_f_p_RangeSizeAndCellContentsAreEqual" 'Name of the function
      .b_prop_rw_SilentError = True 'False will display a message box - you should only do this on entry level
      .s_prop_rw_ErrorMessage = "An execution error occurred." 'A message that properly informs the user and the devs (silent errors will be logged nonetheless)
      .SetCallArgs "No args" 'If the sub takes args put the here like ("sExample:=" & sExample, "lExample:=" & lExample)
   End With
'Fixed, don't change
Try: On Error GoTo Catch


'>>>>>>> Your code here
   Dim sTempMessage As String
   Dim bTempResult As Boolean
   Dim lRow As Long
   Dim lColumn As Long
   
   sTempMessage = "Size and cell contents are equal."
   bTempResult = True

   ' Check if ranges are the same size
   If oRngOne.Rows.Count <> oRngTwo.Rows.Count _
   Or oRngOne.Columns.Count <> oRngTwo.Columns.Count Then
      sTempMessage = "Size is different."
      bTempResult = False
      GoTo Finally
   End If
    
   ' Check if ranges match
   For lRow = 1 To oRngOne.Rows.Count
      For lColumn = 1 To oRngOne.Columns.Count
         If oRngOne.Cells(lRow, lColumn).Value2 <> oRngTwo.Cells(lRow, lColumn).Value2 Then
            bTempResult = False
            sTempMessage = "Cell contents is different."
            GoTo Finally
         End If
      Next lColumn
   Next lRow
    
'End of your code <<<<<<<


'Fixed, don't change
Finally: On Error Resume Next


'>>>>>>> Your code here
   sStatusMessage = sTempMessage
   bAreEqual = bTempResult

'End of your code <<<<<<<


'>>>>>>> Your custom settings here
   If oC_Me.oC_prop_r_Error Is Nothing Then b_f_p_RangeSizeAndCellContentsAreEqual = True 'reports execution as successful to caller
'Fixed, don't change
   Exit Function
HandleError: af_pM_ErrorHandling.af_p_Hook_ErrorHandling_LowerLevel


'>>>>>>> Your code here
   sTempMessage = "An error occurred during processing."


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
End Function

