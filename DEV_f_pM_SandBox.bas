Attribute VB_Name = "DEV_f_pM_SandBox"
' CORE-DEV - do not change, optionally remove when deploying app
'============================================================================================
'   NAME:     DEV_f_pM_SandBox
'============================================================================================
'   Purpose: Sandbox for dev experiments
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
'--------------------------------------------------------------------------------------------
'   BACKLOG
'   ''''''''''''''''''''
'   none
'============================================================================================
Option Explicit
Option Private Module

Private Const s_m_COMPONENT_NAME As String = "DEV_f_pM_SandBox"

' Purpose: tests zero sanitation manually, if the code executes without stopping the tests were successful
' 0.11.0    05.08.2022    gueleh    Initially created
Private Sub mManualTest_RangeArrayProcessorZeroSanitation()
   Dim oC As New f_C_RangeArrayProcessor
   Dim va() As Variant
   DEV_Reset_DEV_f_wks_TestCanvas
   With DEV_f_wks_TestCanvas
      .Range("A1").Value = "ID"
      .Range("B1").Value = "ID2"
      .Range("C1").Value = "Value"
      .Range("A2").Value = "'01"
      .Range("B2").Value = "AgA"
      .Range("C2").Value = "What is AgA?"
      .Columns.AutoFit
      va = .Range("A1").CurrentRegion.Formula
      .Range("A1").CurrentRegion.Formula = va
      Debug.Assert .Range("A2").Value = "1"
      oC.SanitizeLeadingZeroItems va
      .Range("A1").CurrentRegion.Formula = va
      Debug.Assert .Range("A2").Value = "01"
   End With
   DEV_Reset_DEV_f_wks_TestCanvas
End Sub
