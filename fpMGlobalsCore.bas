Attribute VB_Name = "fpMGlobalsCore"
' -------------------------------------------------------------------------------------------
' CORE, do not change
'============================================================================================
'   NAME:     fpMGlobalsCore
'============================================================================================
'   Purpose:  the globals for the core part of the framework
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

Private Const smCOMPONENT_NAME As String = "fpMGlobalsCore"

'determines which mode supposed to be executed
Public Enum efProcessingModes
   efProcessingModeGlobalsOnly
   efProcessingModeAppSpecific
   efProcessingModeAutoCalcOffOnSceenUpdatingOffOn
End Enum

'global class with framework settings, instance created during initialization of globals
Public oCfgFrameworkSettings As fCSettings
'global collection for error handling, errors are added to it and in the end the whole collection is handled based on it
Public oColfgErrors As Collection
'global collection for unit testing, test documentation is added to it and in the end a report is generated based on it
Public oColfgUnitTests As Collection

' Purpose: starts the processing, to be executed at the very begin of the entry level
' 0.1.0    20220709    gueleh    Initially created
Public Sub fStartProcessing( _
   Optional ByVal efProcessingMode As efProcessingModes = 0, _
   Optional ByVal eafProcessingMode As eafProcessingModes = 0)
   
   fInitGlobals
   mfStartProcessingMode efProcessingMode, eafProcessingMode
   
End Sub

' Purpose: ends the processing, to be executed at the very end of the entry level
' 0.1.0    20220709    gueleh    Initially created
Public Sub fEndProcessing( _
   Optional ByVal efProcessingMode As efProcessingModes = 0, _
   Optional ByVal eafProcessingMode As eafProcessingModes = 0)
   
   mfEndProcessingMode efProcessingMode, eafProcessingMode
   
End Sub

' Purpose: initializes the globals that are part of the framework's core
' 0.1.0    20220709    gueleh    Initially created
Public Sub fInitGlobals()
   mResetGlobals
   Set oCfgFrameworkSettings = New fCSettings
   Set oColfgErrors = New Collection
   
   ' Only executed when components are present
   On Error Resume Next
   ' Globals initialization for development contents
   Application.Run sfRunMyProcedure("devfInitGlobals")
End Sub

' Purpose: registers a unit test if tests are supposed to be run, but only if the required modules are in the project
' 0.1.0    20220709    gueleh    Initially created
Public Sub fRegisterUnitTest(ByRef oCParams As fCCallParams)
   On Error Resume Next
   Application.Run sfRunMyProcedure("devfRegisterUnitTest"), oCParams
End Sub

' Purpose: registers an execution error for a unit in a unit test if tests are run, but only if the required modules are in the project
' 0.1.0    20220709    gueleh    Initially created
Public Sub fRegisterExecutionError(ByRef oCParams As fCCallParams)
   On Error Resume Next
   Application.Run sfRunMyProcedure("devfRegisterExecutionError"), oCParams
End Sub

' Purpose: reset the globals which should not retain their value
' 0.1.0    20220709    gueleh    Initially created
Private Sub mResetGlobals()
   Set oColfgUnitTests = Nothing
End Sub

' Purpose: supports several different modes for starting the processing
' 0.1.0    20220709    gueleh    Initially created
Private Sub mfStartProcessingMode( _
   ByVal efProcessingMode As efProcessingModes, _
   ByVal eafProcessingMode As eafProcessingModes)
   
   Select Case efProcessingMode
      Case efProcessingModeAutoCalcOffOnSceenUpdatingOffOn
         With Application
            .ScreenUpdating = False
            .Calculation = xlCalculationManual
         End With
      Case efProcessingModeGlobalsOnly
         'Do nothing except for the required initialization
      Case efProcessingModeAppSpecific
         afStartProcessingMode eafProcessingMode
      Case Else
         'Do nothing except for the required initialization
   End Select
End Sub

' Purpose: supports several different modes for ending the processing
' 0.1.0    20220709    gueleh    Initially created
Private Sub mfEndProcessingMode( _
   ByVal efProcessingMode As efProcessingModes, _
   ByVal eafProcessingMode As eafProcessingModes)
   
   Select Case efProcessingMode
      Case efProcessingModeAutoCalcOffOnSceenUpdatingOffOn
         With Application
            .ScreenUpdating = True
            .Calculation = xlCalculationAutomatic
            .Calculate
         End With
      Case efProcessingModeGlobalsOnly
         'Do nothing except for the required initialization
      Case efProcessingModeAppSpecific
         afEndProcessingMode eafProcessingMode
      Case Else
         'Do nothing except for the required initialization
   End Select
End Sub


