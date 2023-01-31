VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "devfCVersionControlExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' -------------------------------------------------------------------------------------------
' CORE-DEV, do not change
'============================================================================================
'   NAME:     devfCVersionControlExport
'============================================================================================
'   Purpose:  exports all data relevant for convenient version control
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
' 0.14.0    31.01.2023    gueleh    Added support for exporting wks data and reference data for version control
'   0.12.0    16.08.2022    gueleh    Initially created
'--------------------------------------------------------------------------------------------
'   BACKLOG
'   ''''''''''''''''''''
'   none
'============================================================================================
Option Explicit

Private Const s_m_COMPONENT_NAME As String = "devfCVersionControlExport"

' TODO: Export List Object properties to text file

' Purpose: print project reference data to text file for version control
' 0.14.0    31.01.2023    gueleh    Initially created
Public Function bExportReferenceData() As Boolean
'Fixed, don't change
   Dim oC_Me As New fCCallParams
   oC_Me.s_prop_rw_ComponentName = s_m_COMPONENT_NAME
   If oC_f_p_FrameworkSettings.b_prop_rw_ThisIsATestRun Then f_p_RegisterUnitTest oC_Me

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Your custom settings here
   With oC_Me
      .s_prop_rw_ProcedureName = "bExportReferenceData" 'Name of the function
      .b_prop_rw_SilentError = True 'False will display a message box - you should only do this on entry level
      .s_prop_rw_ErrorMessage = "The export of name data failed." 'A message that properly informs the user and the devs (silent errors will be logged nonetheless)
      .SetCallArgs "No args" 'If the sub takes args put the here like ("sExample:=" & sExample, "lExample:=" & lExample)
   End With
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

'Fixed, don't change
Try:
   On Error GoTo Catch

'>>>>>>> Your code here
   Dim oC As VBIDE.Reference
   Dim iFileNumber As Integer
   Dim sFileFullPath As String
   Dim sReference As String
   
   sFileFullPath = ThisWorkbook.Path & "\" & "References.txt"
   On Error Resume Next
   Kill sFileFullPath
   On Error GoTo Catch
   iFileNumber = FreeFile
   
   Open sFileFullPath For Output As iFileNumber
   
   For Each oC In ThisWorkbook.VBProject.References
      With oC
         sReference = .Name & "|" & .Description & "|" _
            & .GUID & "|" & .Major & "|" & .Minor & "|" & .FullPath
      End With
      Print #iFileNumber, sReference
   Next oC

'End of your code <<<<<<<

'Fixed, don't change
Finally:
   On Error Resume Next
'>>>>>>> Your code here
   'everything that must be executed regardless of an error or not
   Close #iFileNumber
'End of your code <<<<<<<

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Your custom settings here
'MAKE SURE TO REPLACE fbTemplateLowerLevel WITH THE NAME OF YOUR FUNCTION
   If oC_Me.oC_prop_r_Error Is Nothing Then bExportReferenceData = True 'reports execution as successful to caller
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


' Purpose: exports names and codenames of worksheets into text file for version control
' 0.14.0    31.01.2023    gueleh    Initially created
Public Function bExportWorksheetNameData() As Boolean
'Fixed, don't change
   Dim oC_Me As New fCCallParams
   oC_Me.s_prop_rw_ComponentName = s_m_COMPONENT_NAME
   If oC_f_p_FrameworkSettings.b_prop_rw_ThisIsATestRun Then f_p_RegisterUnitTest oC_Me

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Your custom settings here
   With oC_Me
      .s_prop_rw_ProcedureName = "bExportWorksheetNameData" 'Name of the function
      .b_prop_rw_SilentError = True 'False will display a message box - you should only do this on entry level
      .s_prop_rw_ErrorMessage = "The export of name data failed." 'A message that properly informs the user and the devs (silent errors will be logged nonetheless)
      .SetCallArgs "No args" 'If the sub takes args put the here like ("sExample:=" & sExample, "lExample:=" & lExample)
   End With
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

'Fixed, don't change
Try:
   On Error GoTo Catch

'>>>>>>> Your code here
   Dim oWks As Worksheet
   Dim iFileNumber As Integer
   Dim sFileFullPath As String
   
   sFileFullPath = ThisWorkbook.Path & "\" & "WorksheetNames.txt"
   On Error Resume Next
   Kill sFileFullPath
   On Error GoTo Catch
   iFileNumber = FreeFile
   
   Open sFileFullPath For Output As iFileNumber
   
   For Each oWks In ThisWorkbook.Worksheets
      Print #iFileNumber, oWks.CodeName, oWks.Name
   Next oWks

'End of your code <<<<<<<

'Fixed, don't change
Finally:
   On Error Resume Next
'>>>>>>> Your code here
   'everything that must be executed regardless of an error or not
   Close #iFileNumber
'End of your code <<<<<<<

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Your custom settings here
'MAKE SURE TO REPLACE fbTemplateLowerLevel WITH THE NAME OF YOUR FUNCTION
   If oC_Me.oC_prop_r_Error Is Nothing Then bExportWorksheetNameData = True 'reports execution as successful to caller
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



' Purpose: exports data of all names to a text file
' 0.12.0    16.08.2022    gueleh    Initially created
Public Function bExportNameData() As Boolean
'Fixed, don't change
   Dim oC_Me As New fCCallParams
   oC_Me.s_prop_rw_ComponentName = s_m_COMPONENT_NAME
   If oC_f_p_FrameworkSettings.b_prop_rw_ThisIsATestRun Then f_p_RegisterUnitTest oC_Me

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Your custom settings here
   With oC_Me
      .s_prop_rw_ProcedureName = "bExportNameData" 'Name of the function
      .b_prop_rw_SilentError = True 'False will display a message box - you should only do this on entry level
      .s_prop_rw_ErrorMessage = "The export of name data failed." 'A message that properly informs the user and the devs (silent errors will be logged nonetheless)
      .SetCallArgs "No args" 'If the sub takes args put the here like ("sExample:=" & sExample, "lExample:=" & lExample)
   End With
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

'Fixed, don't change
Try:
   On Error GoTo Catch

'>>>>>>> Your code here
   Dim oName As Name
   Dim iFileNumber As Integer
   Dim sFileFullPath As String
   
   sFileFullPath = ThisWorkbook.Path & "\" & "Names.txt"
   On Error Resume Next
   Kill sFileFullPath
   On Error GoTo Catch
   iFileNumber = FreeFile
   
   Open sFileFullPath For Output As iFileNumber
   
   For Each oName In ThisWorkbook.Names
      If oName.MacroType <> xlNone Then
         Print #iFileNumber, oName.Name, oName.Category, oName.Parent.Name, oName.Comment
      Else
         Print #iFileNumber, oName.Name, oName.RefersTo, oName.Parent.Name, oName.Comment
      End If
   Next oName

'End of your code <<<<<<<

'Fixed, don't change
Finally:
   On Error Resume Next
'>>>>>>> Your code here
   'everything that must be executed regardless of an error or not
   Close #iFileNumber
'End of your code <<<<<<<

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Your custom settings here
'MAKE SURE TO REPLACE fbTemplateLowerLevel WITH THE NAME OF YOUR FUNCTION
   If oC_Me.oC_prop_r_Error Is Nothing Then bExportNameData = True 'reports execution as successful to caller
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

' Purpose: exports all VB components of types worksheet, module, class module and form to the project folder
' 0.12.0    16.08.2022    gueleh    Initially created
Public Function bExportAllComponents() As Boolean

'Fixed, don't change
   Dim oC_Me As New fCCallParams
   oC_Me.s_prop_rw_ComponentName = s_m_COMPONENT_NAME
   If oC_f_p_FrameworkSettings.b_prop_rw_ThisIsATestRun Then f_p_RegisterUnitTest oC_Me

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Your custom settings here
   With oC_Me
      .s_prop_rw_ProcedureName = "bExportAllComponents" 'Name of the function
      .b_prop_rw_SilentError = True 'False will display a message box - you should only do this on entry level
      .s_prop_rw_ErrorMessage = "The export of components failed." 'A message that properly informs the user and the devs (silent errors will be logged nonetheless)
      .SetCallArgs "No args" 'If the sub takes args put the here like ("sExample:=" & sExample, "lExample:=" & lExample)
   End With
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

'Fixed, don't change
Try:
   On Error GoTo Catch

'>>>>>>> Your code here
   Dim ocmpComponent As VBIDE.VBComponent
   Dim bExport As Boolean
   Dim sFileName As String
   
   For Each ocmpComponent In ThisWorkbook.VBProject.VBComponents
      bExport = True
      sFileName = ocmpComponent.Name
      Select Case ocmpComponent.Type
         Case vbext_ct_ClassModule, 100 '100 is a worksheet
            sFileName = sFileName & ".cls"
         Case vbext_ct_MSForm
            sFileName = sFileName & ".frm"
         Case vbext_ct_StdModule
            sFileName = sFileName & ".bas"
         Case vbext_ct_Document
            bExport = False
      End Select
      
      If bExport Then
         ocmpComponent.Export ThisWorkbook.Path & "\" & sFileName
      End If
      
   Next ocmpComponent
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
   If oC_Me.oC_prop_r_Error Is Nothing Then bExportAllComponents = True 'reports execution as successful to caller
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


