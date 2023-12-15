Attribute VB_Name = "af_pM_Globals"
' APP-SPECIFIC CORE MODULE - you have to migrate app contents manually in case of a template update
'============================================================================================
'   NAME:     af_pM_Globals
'============================================================================================
'   Purpose:  the app-specific globals of the framework
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
'   0.1.0    20220709    gueleh    Initially created
'--------------------------------------------------------------------------------------------
'   BACKLOG
'   ''''''''''''''''''''
'   none
'============================================================================================
Option Explicit
Option Private Module

Private Const s_m_COMPONENT_NAME As String = "af_pM_Globals"

'determines which mode supposed to be executed
Public Enum e_af_p_ProcessingModes
   e_af_p_ProcessingModeGlobalsOnly
'>>>>>>> add your modes here - then modify af_g_StartProcessingMode below to add your code
' for your modes

'<<<<<<<
End Enum

' Purpose: executes the start processing logic as determined by the app-specific case
' "start processing" is what is done as a first step when running any code, which always
' should start in a public entry level module, see the template procedure in f_pM_TemplatesCore
' Template Versions:
' 0.1.0    20220709    gueleh    Initially created
Public Sub af_p_StartProcessingMode(ByVal e_arg_ProcessingMode As e_af_p_ProcessingModes)
   Select Case e_arg_ProcessingMode
      Case e_af_p_ProcessingModeGlobalsOnly
         'Do nothing except for the required initialization
'>>>>>>> Your cases here
      'Case e_af_p_ProcessingModeMyFineMode
         'My fine code for this processing mode
         
'<<<<<<<
      Case Else
         'Do nothing except for the required initialization
   End Select
End Sub

' Purpose: executes the start processing logic as determined by the app-specific case
' "end processing" is what is done at the very end of the entry level procedure, i.e. it
' is the last code executed before code execution ends
' Template Versions:
' 0.1.0    20220709    gueleh    Initially created
Public Sub af_p_EndProcessingMode(ByVal e_arg_ProcessingMode As e_af_p_ProcessingModes)
   Select Case e_arg_ProcessingMode
      Case e_af_p_ProcessingModeGlobalsOnly
         'Do nothing
'>>>>>>> Your cases here
         
'<<<<<<<
      Case Else
         'Do nothing
   End Select
End Sub

' Purpose: builds and returns collection with f_C_SettingsSheet instances
'  for the provided worksheets
' 1.3.0    18.10.2023    gueleh    Initially created
Public Function oCol_af_p_SettingsSheets() As Collection
   
   Const lROW_START As Long = 3
   Const lCOL_ID As Long = 3
   Const lCOL_NAME As Long = 1
   Const lCOL_VALUE As Long = 2
   
   Dim oCol As New Collection
   Dim oC As New f_C_SettingsSheet
   Dim oColWks As New Collection
   Dim oWks As Worksheet
   
   On Error GoTo Catch
   
   With oColWks
      .Add f_wks_Settings
      .Add af_wks_Settings
      .Add a_wks_Settings
      .Add a_wks_VersionControlRanges
   End With
   
   For Each oWks In oColWks
      Set oC = New f_C_SettingsSheet
         If Not _
      oC.bConstruct(oWks, lROW_START, lCOL_ID, lCOL_NAME, lCOL_VALUE) _
         Then Err.Raise _
            e_f_p_HandledError_GeneralError, , _
            s_f_p_HandledErrorDescription(e_f_p_HandledError_GeneralError)
      oCol.Add oC
   Next oWks
      
'>>>>>>> Your settings sheets here
' (they have to fulfill the contract, refer to f_C_SettingsSheet for guidance)
   
'<<<<<<<
   Set oCol_af_p_SettingsSheets = oCol
   Exit Function
   
Catch:
   Set oCol_af_p_SettingsSheets = Nothing

End Function

