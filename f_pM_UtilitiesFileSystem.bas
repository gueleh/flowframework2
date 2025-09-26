Attribute VB_Name = "f_pM_UtilitiesFileSystem"
' -------------------------------------------------------------------------------------------
' CORE, do not change
'============================================================================================
'   NAME:     f_pM_UtilitiesFileSystem
'============================================================================================
'   Purpose:  utilities for file system handling
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
'   1.15.0    15.08.2024    gueleh    Initially created
'--------------------------------------------------------------------------------------------
'   BACKLOG
'   ''''''''''''''''''''
'   none
'============================================================================================
Option Explicit
Option Private Module

Private Const s_m_COMPONENT_NAME As String = "f_pM_UtilitiesFileSystem"

' Purpose: opens an Excel workbook based on provided path and options and returns it in the provided workbook object
Public Function b_f_p_GetWorkbookFromFullName( _
    ByRef oWkb As Workbook, _
    ByVal sFullName As String, _
    Optional ByVal bUpdateLinks As Boolean = False, _
    Optional ByVal bReadOnly As Boolean = False _
) As Boolean

'Fixed, don't change
   Dim oC_Me As New f_C_CallParams: oC_Me.s_prop_rw_ComponentName = s_m_COMPONENT_NAME: If oC_f_p_FrameworkSettings.b_prop_rw_ThisIsATestRun Then f_p_RegisterUnitTest oC_Me
'>>>>>>> Your custom settings here
   With oC_Me
      .s_prop_rw_ProcedureName = "b_f_p_GetWorkbookFromFullName" 'Name of the function
      .b_prop_rw_SilentError = True 'False will display a message box - you should only do this on entry level
      .s_prop_rw_ErrorMessage = "Opening the workbook failed." 'A message that properly informs the user and the devs (silent errors will be logged nonetheless)
      .SetCallArgs "sFullName:=" & sFullName, "bUpdateLinks:=" & bUpdateLinks, "bReadOnly:=" & bReadOnly 'If the sub takes args put the here like ("sExample:=" & sExample, "lExample:=" & lExample)
   End With

'Fixed, don't change
Try: On Error GoTo Catch
'>>>>>>> Your code here
    Dim vUpdateLinks As Variant
    If bUpdateLinks Then
        vUpdateLinks = 1
    Else
        vUpdateLinks = 0
    End If
    Set oWkb = Workbooks.Open(sFullName, vUpdateLinks, bReadOnly)

'End of your code <<<<<<<

'Fixed, don't change
Finally: On Error Resume Next
'>>>>>>> Your code here

'End of your code <<<<<<<
'>>>>>>> Your custom settings here
   If oC_Me.oC_prop_r_Error Is Nothing Then b_f_p_GetWorkbookFromFullName = True 'reports execution as successful to caller
'Fixed, don't change
   Exit Function

HandleError: af_pM_ErrorHandling.af_p_Hook_ErrorHandling_LowerLevel
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
End Function

'Purpose: provides workbook object and CodeName of worksheet to be found and returns worksheet object if found, otherwise an execution error
Public Function b_f_p_GetWorksheetFromCodeName( _
   ByRef oWks As Worksheet, _
   ByVal sCodeName As String, _
   ByRef oWkb As Workbook _
) As Boolean

'Fixed, don't change
   Dim oC_Me As New f_C_CallParams: oC_Me.s_prop_rw_ComponentName = s_m_COMPONENT_NAME: If oC_f_p_FrameworkSettings.b_prop_rw_ThisIsATestRun Then f_p_RegisterUnitTest oC_Me
'>>>>>>> Your custom settings here
   With oC_Me
      .s_prop_rw_ProcedureName = "b_f_p_GetWorksheetFromCodeName" 'Name of the function
      .b_prop_rw_SilentError = True 'False will display a message box - you should only do this on entry level
      .s_prop_rw_ErrorMessage = "Worksheet not found." 'A message that properly informs the user and the devs (silent errors will be logged nonetheless)
      .SetCallArgs "sCodeName:=" & sCodeName, "oWkb.Name:=" & oWkb.name 'If the sub takes args put the here like ("sExample:=" & sExample, "lExample:=" & lExample)
   End With

'Fixed, don't change
Try: On Error GoTo Catch
'>>>>>>> Your code here
   Dim oWksCandidate As Worksheet
   Dim bFound As Boolean
   For Each oWksCandidate In oWkb.Worksheets
      If oWksCandidate.CodeName = sCodeName Then
         bFound = True
         Set oWks = oWksCandidate
         Exit For
      End If
   Next oWksCandidate
   If Not bFound Then Err.Raise e_f_p_HandledError_GeneralError, , s_f_p_HandledErrorDescription(e_f_p_HandledError_GeneralError)
   
'End of your code <<<<<<<

'Fixed, don't change
Finally: On Error Resume Next
'>>>>>>> Your code here

'End of your code <<<<<<<
'>>>>>>> Your custom settings here
   If oC_Me.oC_prop_r_Error Is Nothing Then b_f_p_GetWorksheetFromCodeName = True 'reports execution as successful to caller
'Fixed, don't change
   Exit Function

HandleError: af_pM_ErrorHandling.af_p_Hook_ErrorHandling_LowerLevel
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
End Function
