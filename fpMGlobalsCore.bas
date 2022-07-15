Attribute VB_Name = "fmGlobalsCore"
' -------------------------------------------------------------------------------------------
' CORE, do not change
'============================================================================================
'   NAME:     fmGlobalsCore
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

Private Const msCOMPONENT_NAME As String = "fmGlobalsCore"

'determines which mode supposed to be executed
Public Enum fProcessingModes
   fProcessingModeGlobalsOnly
   fProcessingModeAppSpecific
   fProcessingModeAutoCalcOffOnSceenUpdatingOffOn
End Enum

'global class with framework settings, instance created during initialization of globals
Public fgclsFrameworkSettings As fclsSettings
'global collection for error handling, errors are added to it and in the end the whole collection is handled based on it
Public fgcolErrors As Collection
'global collection for unit testing, test documentation is added to it and in the end a report is generated based on it
Public fgcolUnitTests As Collection

' Purpose: starts the processing, to be executed at the very begin of the entry level
' 0.1.0    20220709    gueleh    Initially created
Public Sub fStartProcessing( _
   Optional ByVal feProcessingMode As fProcessingModes = 0, _
   Optional ByVal afeProcessingMode As afProcessingModes = 0)
   
   fInitGlobals
   mfStartProcessingMode feProcessingMode, afeProcessingMode
   
End Sub

' Purpose: ends the processing, to be executed at the very end of the entry level
' 0.1.0    20220709    gueleh    Initially created
Public Sub fEndProcessing( _
   Optional ByVal feProcessingMode As fProcessingModes = 0, _
   Optional ByVal afeProcessingMode As afProcessingModes = 0)
   
   mfEndProcessingMode feProcessingMode, afeProcessingMode
   
End Sub

' Purpose: initializes the globals that are part of the framework's core
' 0.1.0    20220709    gueleh    Initially created
Public Sub fInitGlobals()
   mResetGlobals
   Set fgclsFrameworkSettings = New fclsSettings
   Set fgcolErrors = New Collection
   
   ' Only executed when components are present
   On Error Resume Next
   ' Globals initialization for development contents
   Application.Run fsRunMyProcedure("devfInitGlobals")
End Sub

' Purpose: registers a unit test if tests are supposed to be run, but only if the required modules are in the project
' 0.1.0    20220709    gueleh    Initially created
Public Sub fRegisterUnitTest(ByRef clsParams As fclsCallParams)
   On Error Resume Next
   Application.Run fsRunMyProcedure("devfRegisterUnitTest"), clsParams
End Sub

' Purpose: registers an execution error for a unit in a unit test if tests are run, but only if the required modules are in the project
' 0.1.0    20220709    gueleh    Initially created
Public Sub fRegisterExecutionError(ByRef clsParams As fclsCallParams)
   On Error Resume Next
   Application.Run fsRunMyProcedure("devfRegisterExecutionError"), clsParams
End Sub

' Purpose: reset the globals which should not retain their value
' 0.1.0    20220709    gueleh    Initially created
Private Sub mResetGlobals()
   Set fgcolUnitTests = Nothing
End Sub

' Purpose: supports several different modes for starting the processing
' 0.1.0    20220709    gueleh    Initially created
Private Sub mfStartProcessingMode(ByVal feProcessingMode As fProcessingModes, ByVal afeProcessingMode As afProcessingModes)
   Select Case feProcessingMode
      Case fProcessingModeAutoCalcOffOnSceenUpdatingOffOn
         With Application
            .ScreenUpdating = False
            .Calculation = xlCalculationManual
         End With
      Case fProcessingModeGlobalsOnly
         'Do nothing except for the required initialization
      Case fProcessingModeAppSpecific
         afStartProcessingMode afeProcessingMode
      Case Else
         'Do nothing except for the required initialization
   End Select
End Sub

' Purpose: supports several different modes for ending the processing
' 0.1.0    20220709    gueleh    Initially created
Private Sub mfEndProcessingMode(ByVal feProcessingMode As fProcessingModes, ByVal afeProcessingMode As afProcessingModes)
   Select Case feProcessingMode
      Case fProcessingModeAutoCalcOffOnSceenUpdatingOffOn
         With Application
            .ScreenUpdating = True
            .Calculation = xlCalculationAutomatic
            .Calculate
         End With
      Case fProcessingModeGlobalsOnly
         'Do nothing except for the required initialization
      Case fProcessingModeAppSpecific
         afEndProcessingMode afeProcessingMode
      Case Else
         'Do nothing except for the required initialization
   End Select
End Sub


