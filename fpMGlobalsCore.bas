Attribute VB_Name = "fpMGlobalsCore"
' -------------------------------------------------------------------------------------------
' CORE, do not change
'============================================================================================
'   NAME:     fpMGlobalsCore
'============================================================================================
'   Purpose:  the globals for the core part of the framework
'   Access:   Private
'   Type:     Module
'   Author:   G�nther Lehner
'   Contact:  guenther.lehner@protonmail.com
'   GitHubID: gueleh
'   Required:
'   Usage:
'--------------------------------------------------------------------------------------------
'   VERSION HISTORY
'   Version    Date    Developer    Changes
'   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 0.9.0    03.08.2022    gueleh    Added const declaration
'   0.1.0    20220709    gueleh    Initially created
'--------------------------------------------------------------------------------------------
'   BACKLOG
'   ''''''''''''''''''''
'   none
'============================================================================================
Option Explicit
Option Private Module

Private Const s_m_COMPONENT_NAME As String = "fpMGlobalsCore"

Public Const s_f_p_SPLIT_SEED_SEPARATOR As String = ","
'TODO: Also add identifier for app tech sheets as soon as it is required
Public Const s_f_p_split_seed_TECH_WKS_IDENTIFIERS As String = "devfwks,devafwks,devawks,fwks,afwks"

'determines which mode is supposed to be executed
Public Enum e_f_p_ProcessingModes
   e_f_p_ProcessingMode_GlobalsOnly
   e_f_p_ProcessingMode_AppSpecific
   e_f_p_ProcessingMode_AutoCalcOffOnSceenUpdatingOffOn
End Enum

'global class with framework settings, instance created during initialization of globals
Public oC_f_p_FrameworkSettings As fCSettings
'global collection for error handling, errors are added to it and in the end the whole collection is handled based on it
Public oCol_f_p_Errors As Collection
'global collection for unit testing, test documentation is added to it and in the end a report is generated based on it
Public oCol_f_p_UnitTests As Collection

' Purpose: starts the processing, to be executed at the very begin of the entry level
' 0.1.0    20220709    gueleh    Initially created
Public Sub f_p_StartProcessing( _
   Optional ByVal e_f_ProcessingMode As e_f_p_ProcessingModes = 0, _
   Optional ByVal e_af_ProcessingMode As e_af_p_ProcessingModes = 0)
   
   f_p_InitGlobals
   f_m_StartProcessingMode e_f_ProcessingMode, e_af_ProcessingMode
   
End Sub

' Purpose: ends the processing, to be executed at the very end of the entry level
' 0.1.0    20220709    gueleh    Initially created
Public Sub f_p_EndProcessing( _
   Optional ByVal e_f_ProcessingMode As e_f_p_ProcessingModes = 0, _
   Optional ByVal e_af_ProcessingMode As e_af_p_ProcessingModes = 0)
   
   f_m_EndProcessingMode e_f_ProcessingMode, e_af_ProcessingMode
   
End Sub

' Purpose: initializes the globals that are part of the framework's core
' 0.1.0    20220709    gueleh    Initially created
Public Sub f_p_InitGlobals()
   f_m_ResetGlobals
   Set oC_f_p_FrameworkSettings = New fCSettings
   Set oCol_f_p_Errors = New Collection
   
   ' Only executed when components are present
   On Error Resume Next
   ' Globals initialization for development contents
   Application.Run s_f_p_MyProcedureName("DEV_f_p_InitGlobals")
End Sub

' Purpose: registers a unit test if tests are supposed to be run, but only if the required modules are in the project
' 0.1.0    20220709    gueleh    Initially created
Public Sub f_p_RegisterUnitTest(ByRef oC_f_Params As fCCallParams)
   On Error Resume Next
   Application.Run s_f_p_MyProcedureName("DEV_f_p_RegisterUnitTest"), oC_f_Params
End Sub

' Purpose: registers an execution error for a unit in a unit test if tests are run, but only if the required modules are in the project
' 0.1.0    20220709    gueleh    Initially created
Public Sub f_p_RegisterExecutionError(ByRef oC_f_Params As fCCallParams)
   On Error Resume Next
   Application.Run s_f_p_MyProcedureName("DEV_f_p_RegisterExecutionError"), oC_f_Params
End Sub

' Purpose: reset the globals which should not retain their value
' 0.1.0    20220709    gueleh    Initially created
Private Sub f_m_ResetGlobals()
   Set oCol_f_p_UnitTests = Nothing
End Sub

' Purpose: supports several different modes for starting the processing
' 0.1.0    20220709    gueleh    Initially created
Private Sub f_m_StartProcessingMode( _
   ByVal e_f_ProcessingMode As e_f_p_ProcessingModes, _
   ByVal e_af_ProcessingMode As e_af_p_ProcessingModes)
   
   Select Case e_f_ProcessingMode
      Case e_f_p_ProcessingMode_AutoCalcOffOnSceenUpdatingOffOn
         With Application
            .ScreenUpdating = False
            .Calculation = xlCalculationManual
         End With
      Case e_f_p_ProcessingMode_GlobalsOnly
         'Do nothing except for the required initialization
      Case e_f_p_ProcessingMode_AppSpecific
         af_p_StartProcessingMode e_af_ProcessingMode
      Case Else
         'Do nothing except for the required initialization
   End Select
End Sub

' Purpose: supports several different modes for ending the processing
' 0.1.0    20220709    gueleh    Initially created
Private Sub f_m_EndProcessingMode( _
   ByVal e_f_ProcessingMode As e_f_p_ProcessingModes, _
   ByVal e_af_ProcessingMode As e_af_p_ProcessingModes)
   
   Select Case e_f_ProcessingMode
      Case e_f_p_ProcessingMode_AutoCalcOffOnSceenUpdatingOffOn
         With Application
            .ScreenUpdating = True
            .Calculation = xlCalculationAutomatic
            .Calculate
         End With
      Case e_f_p_ProcessingMode_GlobalsOnly
         'Do nothing except for the required initialization
      Case e_f_p_ProcessingMode_AppSpecific
         af_p_EndProcessingMode e_af_ProcessingMode
      Case Else
         'Do nothing except for the required initialization
   End Select
End Sub


