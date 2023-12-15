VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "f_C_Deploy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' -------------------------------------------------------------------------------------------
' framework
'============================================================================================
'   NAME:     f_C_Deploy
'============================================================================================
'   Purpose:  saves a prod copy of the workbook and removes modules starting with DEV
'   Access:   Public
'   Type:     Class Module
'   Author:   Günther Lehner
'   Contact:  guleh@pm.me
'   GitHubID: gueleh
'   Required:
'   Usage:
'--------------------------------------------------------------------------------------------
'   VERSION HISTORY
'   Version    Date    Developer    Changes
'   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   1.13.0    15.12.2023    gueleh    Initially created
'--------------------------------------------------------------------------------------------
'   BACKLOG
'   ''''''''''''''''''''
'   none
'============================================================================================
Option Explicit

Private Const s_m_COMPONENT_NAME As String = "f_C_Deploy"

Public Function bSaveAsProdAndRemoveDEVModules() As Boolean

'Fixed, don't change
   Dim oC_Me As New f_C_CallParams
   oC_Me.s_prop_rw_ComponentName = s_m_COMPONENT_NAME
   If oC_f_p_FrameworkSettings.b_prop_rw_ThisIsATestRun Then f_p_RegisterUnitTest oC_Me

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Your custom settings here
   With oC_Me
      .s_prop_rw_ProcedureName = "bSaveAsProdAndRemoveDEVModules" 'Name of the function
      .b_prop_rw_SilentError = True 'False will display a message box - you should only do this on entry level
      .s_prop_rw_ErrorMessage = "The removal of components failed." 'A message that properly informs the user and the devs (silent errors will be logged nonetheless)
      .SetCallArgs "No args" 'If the sub takes args put the here like ("sExample:=" & sExample, "lExample:=" & lExample)
   End With
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

'Fixed, don't change
Try:
   On Error GoTo Catch

'>>>>>>> Your code here
   Dim ocmpComponent As VBIDE.VBComponent
   Dim oColComponentsToRemove As New Collection
   
   ThisWorkbook.SaveAs "PROD-" & ThisWorkbook.Name
   
   If Not Left(ThisWorkbook.Name, 5) = "PROD-" Then
      MsgBox "Saving production copy failed, deployment cancelled.", vbInformation
      GoTo Finally
   End If
   
   For Each ocmpComponent In ThisWorkbook.VBProject.VBComponents
      Select Case ocmpComponent.Type
         Case vbext_ct_ClassModule, vbext_ct_MSForm, vbext_ct_StdModule
            If Left(ocmpComponent.Name, 3) = "DEV" Then
               oColComponentsToRemove.Add ocmpComponent
            End If
      End Select
   Next ocmpComponent
   
   If oColComponentsToRemove.Count > 0 Then
      For Each ocmpComponent In oColComponentsToRemove
         ThisWorkbook.VBProject.VBComponents.Remove ocmpComponent
      Next ocmpComponent
   End If
'End of your code <<<<<<<

'Fixed, don't change
Finally:
   On Error Resume Next

'>>>>>>> Your code here
   'everything that must be executed regardless of an error or not
'End of your code <<<<<<<

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Your custom settings here
'MAKE SURE TO REPLACE fbTemplateLowerLevel WITH THE NAME OF YOUR FUNCTION
   If oC_Me.oC_prop_r_Error Is Nothing Then bSaveAsProdAndRemoveDEVModules = True 'reports execution as successful to caller
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

'Fixed, don't change
   Exit Function
Catch:
   If oC_Me.oC_prop_r_Error Is Nothing _
   Then f_p_RegisterError oC_Me, Err.Number, Err.Description
   If oC_f_p_FrameworkSettings.b_prop_rw_ThisIsATestRun Then f_p_RegisterExecutionError oC_Me
   If oC_f_p_FrameworkSettings.b_prop_r_DebugModeIsOn And Not oC_Me.b_prop_rw_ResumedOnce Then
      oC_Me.b_prop_rw_ResumedOnce = True: Stop: Resume
   Else
      f_p_HandleError oC_Me: Resume Finally
   End If

End Function


