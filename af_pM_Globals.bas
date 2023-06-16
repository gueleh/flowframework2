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
Public Sub af_p_StartProcessingMode(ByVal eafProcessingMode As e_af_p_ProcessingModes)
   Select Case eafProcessingMode
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
Public Sub af_p_EndProcessingMode(ByVal eafProcessingMode As e_af_p_ProcessingModes)
   Select Case eafProcessingMode
      Case e_af_p_ProcessingModeGlobalsOnly
         'Do nothing
'>>>>>>> Your cases here
         
'<<<<<<<
      Case Else
         'Do nothing
   End Select
End Sub

