VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DEV_f_C_VersionControlExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' CORE-DEV - do not change, optionally remove when deploying app
'============================================================================================
'   NAME:     DEV_f_C_VersionControlExport
'============================================================================================
'   Purpose:  exports all data relevant for convenient version control
'   Access:   Private
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

Private Const s_m_COMPONENT_NAME As String = "DEV_f_C_VersionControlExport"

Public Function bExportRangeContentData() As Boolean
'Fixed, don't change
   Dim oC_Me As New f_C_CallParams
   oC_Me.s_prop_rw_ComponentName = s_m_COMPONENT_NAME
   If oC_f_p_FrameworkSettings.b_prop_rw_ThisIsATestRun Then f_p_RegisterUnitTest oC_Me

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Your custom settings here
   With oC_Me
      .s_prop_rw_ProcedureName = "bExportRangeContentData" 'Name of the function
      .b_prop_rw_SilentError = True 'False will display a message box - you should only do this on entry level
      .s_prop_rw_ErrorMessage = "The export of range content data failed." 'A message that properly informs the user and the devs (silent errors will be logged nonetheless)
      .SetCallArgs "No args" 'If the sub takes args put the here like ("sExample:=" & sExample, "lExample:=" & lExample)
   End With
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

'Fixed, don't change
Try:
   On Error GoTo Catch

'>>>>>>> Your code here
   Const lROW_START As Long = 3
   Const lCOL_ID As Long = 3
   Const lCOL_NAME As Long = 1
   Const lCOL_VALUE As Long = 2
   
   Dim oWks As Worksheet
   Dim iFileNumber As Integer
   Dim sFileFullPath As String
   Dim oCSettingsSheet As f_C_SettingsSheet
   Dim oColSettings As Collection
   Dim oCSetting As f_C_Setting
   
   Dim oCRanges As New DEV_f_C_VersionControlRanges
   Dim oCRange As DEV_f_C_VersionControlRange
   Dim oColRanges As Collection
   Dim lCountRanges As Long
   Dim oRngCell As Range
   Dim lCountCells As Long
   Dim oWksParent As Worksheet
   
   oCRanges.Construct a_wks_VersionControlRanges, lROW_START, lCOL_NAME, lCOL_ID
   
      If Not _
   oCRanges.bBuildCollectionWithRanges() _
      Then Err.Raise _
         e_f_p_HandledError_ExecutionOfLowerLevelFunction, , _
         s_f_p_HandledErrorDescription(e_f_p_HandledError_ExecutionOfLowerLevelFunction)
   
   Set oColRanges = oCRanges.oCol_prop_r_VersionControlRanges
      
   If oColRanges.Count = 0 Then GoTo Finally
      
   
   sFileFullPath = ThisWorkbook.Path & "\" & "VersionControlledRangeContent.json"
   On Error Resume Next
   Kill sFileFullPath
   On Error GoTo Catch
   iFileNumber = FreeFile
   
   Open sFileFullPath For Output As iFileNumber
   
   For Each oCRange In oColRanges
      lCountRanges = lCountRanges + 1

      Set oWksParent = oCRange.oRng_prop_r_TheRange.Parent
      Print #iFileNumber, "{"
      Print #iFileNumber, vbTab & """" & oCRange.s_prop_r_NameOfRange & """: ["
      Print #iFileNumber, vbTab & vbTab & "{"
      Print #iFileNumber, vbTab & vbTab & vbTab & """name"": """ & oCRange.s_prop_r_Name & ""","
      Print #iFileNumber, vbTab & vbTab & vbTab & """worksheetCodeName"": """ & oWksParent.CodeName & ""","
      Print #iFileNumber, vbTab & vbTab & vbTab & """rowsCount"": """ & oCRange.oRng_prop_r_TheRange.Rows.CountLarge & ""","
      Print #iFileNumber, vbTab & vbTab & vbTab & """columnsCount"": """ & oCRange.oRng_prop_r_TheRange.Columns.CountLarge & ""","
      Print #iFileNumber, vbTab & vbTab & vbTab & """anchorRow"": """ & oCRange.oRng_prop_r_TheRange.Row & ""","
      Print #iFileNumber, vbTab & vbTab & vbTab & """anchorColum"": """ & oCRange.oRng_prop_r_TheRange.Column & ""","
      Print #iFileNumber, vbTab & vbTab & vbTab & """cellContents"": ["
      For Each oRngCell In oCRange.oRng_prop_r_TheRange
         lCountCells = lCountCells + 1
         Print #iFileNumber, vbTab & vbTab & vbTab & vbTab & "{"
         Print #iFileNumber, vbTab & vbTab & vbTab & vbTab & vbTab & """rowNumber"": """ & oRngCell.Row & ""","
         Print #iFileNumber, vbTab & vbTab & vbTab & vbTab & vbTab & """columnNumber"": """ & oRngCell.Column & ""","
         Print #iFileNumber, vbTab & vbTab & vbTab & vbTab & vbTab & """value2"": """ & oRngCell.Value2 & ""","
         Print #iFileNumber, vbTab & vbTab & vbTab & vbTab & vbTab & """formulaLocal"": """ & Replace$(oRngCell.FormulaLocal, """", "\""") & """"
         If lCountCells = oCRange.oRng_prop_r_TheRange.CountLarge Then
            Print #iFileNumber, vbTab & vbTab & vbTab & vbTab & "}"
            Print #iFileNumber, vbTab & vbTab & vbTab & "]"
         Else
            Print #iFileNumber, vbTab & vbTab & vbTab & vbTab & "},"
         End If
      Next oRngCell
      Print #iFileNumber, vbTab & vbTab & "}"
      If lCountRanges = oColRanges.Count Then
         Print #iFileNumber, vbTab & "]"
      Else
         Print #iFileNumber, vbTab & "],"
      End If
   Next oCRange
   Print #iFileNumber, "}"
   
   Close #iFileNumber

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
   If oC_Me.oC_prop_r_Error Is Nothing Then bExportRangeContentData = True 'reports execution as successful to caller
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


' Purpose: exports data from settings sheets
' 1.4.0    19.10.2023    gueleh    changed export from text to json
' 1.3.0    18.10.2023    gueleh    Initially created
Public Function bExportSettingsSheetData() As Boolean
'Fixed, don't change
   Dim oC_Me As New f_C_CallParams
   oC_Me.s_prop_rw_ComponentName = s_m_COMPONENT_NAME
   If oC_f_p_FrameworkSettings.b_prop_rw_ThisIsATestRun Then f_p_RegisterUnitTest oC_Me

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Your custom settings here
   With oC_Me
      .s_prop_rw_ProcedureName = "bExportSettingsSheetData" 'Name of the function
      .b_prop_rw_SilentError = True 'False will display a message box - you should only do this on entry level
      .s_prop_rw_ErrorMessage = "The export of settings sheet data failed." 'A message that properly informs the user and the devs (silent errors will be logged nonetheless)
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
   Dim oColSheets As Collection
   Dim oCSettingsSheet As f_C_SettingsSheet
   Dim oColSettings As Collection
   Dim oCSetting As f_C_Setting
   Dim lSettingsCount As Long
   
   Set oColSheets = oCol_af_p_SettingsSheets()
   
   If oColSheets Is Nothing Then
      MsgBox "Collection with settings sheets could not be build. Creation of version control data for settings sheets skipped.", vbExclamation
      GoTo Finally
   End If
   
   For Each oCSettingsSheet In oColSheets
      Set oColSettings = New Collection
         If Not _
      oCSettingsSheet.bGetSettingsFromSettingsSheet(oColSettings) _
         Then Err.Raise _
            e_f_p_HandledError_ExecutionOfLowerLevelFunction, , _
            s_f_p_HandledErrorDescription(e_f_p_HandledError_ExecutionOfLowerLevelFunction)
      sFileFullPath = ThisWorkbook.Path & "\" & "SettingsSheet-" _
         & oCSettingsSheet.oWks_prop_r_SettingsSheet.CodeName & ".json"
      On Error Resume Next
      Kill sFileFullPath
      On Error GoTo Catch
      iFileNumber = FreeFile
      
      Open sFileFullPath For Output As iFileNumber
      
      Print #iFileNumber, "{"
      Print #iFileNumber, vbTab & """" & oCSettingsSheet.oWks_prop_r_SettingsSheet.CodeName & """: ["
      For Each oCSetting In oColSettings
         lSettingsCount = lSettingsCount + 1
         Print #iFileNumber, vbTab & vbTab & "{"
         Print #iFileNumber, vbTab & vbTab & """id"": """ & oCSetting.sID & ""","
         Print #iFileNumber, vbTab & vbTab & """name"": """ & oCSetting.sName & ""","
         Print #iFileNumber, vbTab & vbTab & """value2"": """ & oCSetting.sValue & ""","
         Print #iFileNumber, vbTab & vbTab & """formulaLocal"": """ & Replace$(oCSetting.sFormulaLocal, """", "\""") & """"
         If lSettingsCount = oColSettings.Count Then
            Print #iFileNumber, vbTab & vbTab & "}"
         Else
            Print #iFileNumber, vbTab & vbTab & "},"
         End If
      Next oCSetting
      Print #iFileNumber, vbTab & "]"
      Print #iFileNumber, "}"
      
      Close #iFileNumber
      
   Next oCSettingsSheet

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
   If oC_Me.oC_prop_r_Error Is Nothing Then bExportSettingsSheetData = True 'reports execution as successful to caller
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



' Purpose: print project reference data to text file for version control
' 1.4.0    19.10.2023    gueleh    changed export from text to json
' 0.14.0    31.01.2023    gueleh    Initially created
Public Function bExportReferenceData() As Boolean
'Fixed, don't change
   Dim oC_Me As New f_C_CallParams
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
   Dim lReferenceCount As Long
   
   sFileFullPath = ThisWorkbook.Path & "\" & "References.json"
   On Error Resume Next
   Kill sFileFullPath
   On Error GoTo Catch
   iFileNumber = FreeFile
   
   Open sFileFullPath For Output As iFileNumber
   
   Print #iFileNumber, "{"
   Print #iFileNumber, vbTab & """references"": ["
   For Each oC In ThisWorkbook.VBProject.References
      lReferenceCount = lReferenceCount + 1
      Print #iFileNumber, vbTab & vbTab & "{"
      Print #iFileNumber, vbTab & vbTab & """name"": """ & oC.Name & ""","
      Print #iFileNumber, vbTab & vbTab & """description"": """ & oC.Description & ""","
      Print #iFileNumber, vbTab & vbTab & """type"": """ & oC.Type & ""","
      Print #iFileNumber, vbTab & vbTab & """builtIn"": """ & oC.BuiltIn & ""","
      Print #iFileNumber, vbTab & vbTab & """guid"": """ & oC.GUID & ""","
      Print #iFileNumber, vbTab & vbTab & """isBroken"": """ & oC.IsBroken & ""","
      Print #iFileNumber, vbTab & vbTab & """major"": """ & oC.Major & ""","
      Print #iFileNumber, vbTab & vbTab & """minor"": """ & oC.Minor & ""","
      Print #iFileNumber, vbTab & vbTab & """fullPath"": """ & Replace$(oC.FullPath, "\", "\\") & """"
      If lReferenceCount = ThisWorkbook.VBProject.References.Count Then
         Print #iFileNumber, vbTab & vbTab & "}"
      Else
         Print #iFileNumber, vbTab & vbTab & "},"
      End If
   Next oC
   Print #iFileNumber, vbTab & "]"
   Print #iFileNumber, "}"

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
' 1.4.0    19.10.2023    gueleh    changed export from text to json
' 0.14.0    31.01.2023    gueleh    Initially created
Public Function bExportWorksheetNameData() As Boolean
'Fixed, don't change
   Dim oC_Me As New f_C_CallParams
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
   Dim lCount As Long
   
   sFileFullPath = ThisWorkbook.Path & "\" & "WorksheetNames.json"
   On Error Resume Next
   Kill sFileFullPath
   On Error GoTo Catch
   iFileNumber = FreeFile
   
   Open sFileFullPath For Output As iFileNumber
   
   Print #iFileNumber, "{"
   Print #iFileNumber, vbTab & """worksheets"": ["
   For Each oWks In ThisWorkbook.Worksheets
      lCount = lCount + 1
      Print #iFileNumber, vbTab & vbTab & "{"
      Print #iFileNumber, vbTab & vbTab & """codeName"": """ & oWks.CodeName & ""","
      Print #iFileNumber, vbTab & vbTab & """name"": """ & oWks.Name & ""","
      Print #iFileNumber, vbTab & vbTab & """usedRangeRowsCountLarge"": """ & oWks.UsedRange.Rows.CountLarge & ""","
      Print #iFileNumber, vbTab & vbTab & """usedRangeColumnsCountLarge"": """ & oWks.UsedRange.Columns.CountLarge & """"
      If lCount = ThisWorkbook.Worksheets.Count Then
         Print #iFileNumber, vbTab & vbTab & "}"
      Else
         Print #iFileNumber, vbTab & vbTab & "},"
      End If
   Next oWks
   Print #iFileNumber, vbTab & "]"
   Print #iFileNumber, "}"

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
' 1.4.0    19.10.2023    gueleh    changed export from text to json
' 0.12.0    16.08.2022    gueleh    Initially created
Public Function bExportNameData() As Boolean
'Fixed, don't change
   Dim oC_Me As New f_C_CallParams
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
   Dim lCount As Long
   
   sFileFullPath = ThisWorkbook.Path & "\" & "Names.json"
   On Error Resume Next
   Kill sFileFullPath
   On Error GoTo Catch
   iFileNumber = FreeFile
   
   Open sFileFullPath For Output As iFileNumber
   
   Print #iFileNumber, "{"
   Print #iFileNumber, vbTab & """names"": ["
   For Each oName In ThisWorkbook.Names
      lCount = lCount + 1
      Print #iFileNumber, vbTab & vbTab & "{"
      
      Print #iFileNumber, vbTab & vbTab & """name"": """ & oName.Name & ""","
      If oName.MacroType <> xlNone Then
         Print #iFileNumber, vbTab & vbTab & """category"": """ & oName.Category & ""","
      Else
         Print #iFileNumber, vbTab & vbTab & """refersTo"": """ & oName.RefersTo & ""","
      End If
      Print #iFileNumber, vbTab & vbTab & """parentName"": """ & oName.Parent.Name & ""","
      Print #iFileNumber, vbTab & vbTab & """comment"": """ & oName.Comment & ""","
      Print #iFileNumber, vbTab & vbTab & """macroType"": """ & oName.MacroType & """"
      
      If lCount = ThisWorkbook.Names.Count Then
         Print #iFileNumber, vbTab & vbTab & "}"
      Else
         Print #iFileNumber, vbTab & vbTab & "},"
      End If
   Next oName
   Print #iFileNumber, vbTab & "]"
   Print #iFileNumber, "}"

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
   Dim oC_Me As New f_C_CallParams
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


