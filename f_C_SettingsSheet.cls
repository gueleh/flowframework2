VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "f_C_SettingsSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' CORE, do not change or remove
'============================================================================================
'   NAME:     f_C_SettingsSheet
'============================================================================================
'   Purpose:  gets settings from a settings sheet which fulfills the structural contract
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
'   1.3.0    18.10.2023    gueleh    Initially created
'--------------------------------------------------------------------------------------------
'   BACKLOG
'   ''''''''''''''''''''
'   none
'============================================================================================
Option Explicit

Private Const s_m_COMPONENT_NAME As String = "f_C_SettingsSheet"

Private oWks_m_Settings As Worksheet
Private l_m_RowStart As Long
Private l_m_ColID As Long
Private l_m_ColName As Long
Private l_m_ColValue As Long

Public Property Get oWks_prop_r_SettingsSheet() As Worksheet
   Set oWks_prop_r_SettingsSheet = oWks_m_Settings
End Property


' Procedure Name: bConstruct
' Purpose: sets up the parameters needed to work with the settings sheet
' Procedure Kind: Function
' Procedure Access: Public
' Parameter oWks (Worksheet): the worksheet object of a settings sheet
' Parameter lRowStart (Long): the row containing the first setting
' Parameter lColID (Long): the column containing the settings ID, e.g. the defined name
' Parameter lColName (Long): the column containing the name of a setting
' Parameter lColValue (Long): the column containing the value of a setting
' Return Type: Boolean
' Author: Günther Lehner
' Contact:  guleh@pm.me
' GitHubID: gueleh
' Requires: a worksheet fulfilling the contract, i.e. having one setting per row
'  with a unique ID, a name and a value in a dedicated column of the worksheet
' Usage:
'------------------------------------------------------------------------------------
' Version   Date      Developer   Changes
' 1.3.0    18.10.2023    gueleh  Initially created
'------------------------------------------------------------------------------------
' Backlog:
' None
'------------------------------------------------------------------------------------
Public Function bConstruct _
( _
   ByRef oWks As Worksheet, _
   ByVal lRowStart As Long, _
   ByVal lColID As Long, _
   ByVal lColName As Long, _
   ByVal lColValue As Long _
) As Boolean

'Fixed, don't change
   Dim oC_Me As New f_C_CallParams: oC_Me.s_prop_rw_ComponentName = s_m_COMPONENT_NAME: If oC_f_p_FrameworkSettings.b_prop_rw_ThisIsATestRun Then f_p_RegisterUnitTest oC_Me
'>>>>>>> Your custom settings here
   With oC_Me
      .s_prop_rw_ProcedureName = "bConstruct" 'Name of the function
      .b_prop_rw_SilentError = True 'False will display a message box - you should only do this on entry level
      .s_prop_rw_ErrorMessage = "Construction of settings sheet failed." 'A message that properly informs the user and the devs (silent errors will be logged nonetheless)
      .SetCallArgs "oWks.CodeName:=" & oWks.CodeName & ", lRowStart:=" & lRowStart & ", lColID:=" & lColID & ", lColName:=" & lColName & ", lColValue:=" & lColValue 'If the sub takes args put the here like ("sExample:=" & sExample, "lExample:=" & lExample)
   End With

'Fixed, don't change
Try: On Error GoTo Catch

'>>>>>>> Your code here
   Set oWks_m_Settings = oWks
   l_m_RowStart = lRowStart
   l_m_ColID = lColID
   l_m_ColName = lColName
   l_m_ColValue = lColValue
'End of your code <<<<<<<

'Fixed, don't change
Finally: On Error Resume Next

'>>>>>>> Your code here
'End of your code <<<<<<<

'>>>>>>> Your custom settings here
   If oC_Me.oC_prop_r_Error Is Nothing Then bConstruct = True 'reports execution as successful to caller

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

' Procedure Name: bGetSettingsFromSettingsSheet
' Purpose: building and returning a collection for all setting rows in a settings sheet
' Procedure Kind: Function
' Procedure Access: Public
' Parameter oCol_Settings (Collection): collection with f_C_Setting instances, one for each row/setting - comes in empty, is returned
' Return Type: Boolean
' Author: Günther Lehner
' Contact:  guleh@pm.me
' GitHubID: gueleh
' Requires:
' Usage:
'------------------------------------------------------------------------------------
' Version   Date      Developer   Changes
' 1.3.0    18.10.2023    gueleh  Initially created
'------------------------------------------------------------------------------------
' Backlog:
' None
'------------------------------------------------------------------------------------
Public Function bGetSettingsFromSettingsSheet _
( _
   ByRef oCol_Settings As Collection _
) As Boolean

'Fixed, don't change
   Dim oC_Me As New f_C_CallParams: oC_Me.s_prop_rw_ComponentName = s_m_COMPONENT_NAME: If oC_f_p_FrameworkSettings.b_prop_rw_ThisIsATestRun Then f_p_RegisterUnitTest oC_Me
'>>>>>>> Your custom settings here
   With oC_Me
      .s_prop_rw_ProcedureName = "bGetSettingsFromSettingsSheet" 'Name of the function
      .b_prop_rw_SilentError = True 'False will display a message box - you should only do this on entry level
      .s_prop_rw_ErrorMessage = "Getting settings failed." 'A message that properly informs the user and the devs (silent errors will be logged nonetheless)
      .SetCallArgs "No args"
   End With

'Fixed, don't change
Try: On Error GoTo Catch

'>>>>>>> Your code here
   Dim oCSetting As f_C_Setting
   Dim lRow As Long
   Dim sID, sName As String
   
   For lRow = l_m_RowStart To oWks_m_Settings.UsedRange.Rows.Count + l_m_RowStart
      sID = CStr(oWks_m_Settings.Cells(lRow, l_m_ColID).FormulaLocal)
      sName = CStr(oWks_m_Settings.Cells(lRow, l_m_ColName).FormulaLocal)
      If Len(sID) > 0 And Len(sName) > 0 Then
         Set oCSetting = New f_C_Setting
         oCSetting.sID = sID
         oCSetting.sName = sName
         oCSetting.sValue = CStr(oWks_m_Settings.Cells(lRow, l_m_ColValue).Value2)
         oCSetting.sFormulaLocal = CStr(oWks_m_Settings.Cells(lRow, l_m_ColValue).FormulaLocal)
         oCol_Settings.Add oCSetting
      End If
   Next lRow
   
'End of your code <<<<<<<

'Fixed, don't change
Finally: On Error Resume Next

'>>>>>>> Your code here
'End of your code <<<<<<<

'>>>>>>> Your custom settings here
   If oC_Me.oC_prop_r_Error Is Nothing Then bGetSettingsFromSettingsSheet = True 'reports execution as successful to caller

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

