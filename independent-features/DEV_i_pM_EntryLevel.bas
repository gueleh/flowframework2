Attribute VB_Name = "DEV_i_pM_EntryLevel"
'============================================================================================
'   NAME:     DEV_pM_EntryLevel
'============================================================================================
'   Purpose:  entry level procedures related to development
'   Access:   Public
'   Type:     Module
'   Author:   Günther Lehner
'   Contact:  guleh@pm.me
'   GitHubID: gueleh
'   Required:
'   Usage:
'--------------------------------------------------------------------------------------------
'   VERSION HISTORY
'   Version    Date    Developer    Changes
'   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 1.0.0  01.06.2023  gueleh   Carved out from framework and removed dependency
'--------------------------------------------------------------------------------------------
'   BACKLOG
'   ''''''''''''''''''''
'   none
'============================================================================================
Option Explicit
Option Private Module

' Purpose: exports all files required for version control to the project folder
' 1.0.0  01.06.2023  gueleh   Carved out from framework and removed dependency
Public Sub DEV_ExportDataForVersionControl()
   
   Const lEXECUTION_ERROR_NUMBER As Long = 9999
   Const sEXECUTION_ERROR_TEXT As String = "Execution of lower level function failed," _
      & "please look into direct window for more information."
   
   Dim bResumedOnce As Boolean
   
Try:
   On Error GoTo Catch
   
   Dim oC_VersionControlExport As New DEV_i_C_VersionControlExport
   
   With Application
      .ScreenUpdating = False
      .Calculation = xlCalculationManual
      .EnableEvents = False
   End With
   
      If Not _
   oC_VersionControlExport.bExportAllComponents() _
      Then Err.Raise lEXECUTION_ERROR_NUMBER, , sEXECUTION_ERROR_TEXT
   
      If Not _
   oC_VersionControlExport.bExportNameData() _
      Then Err.Raise lEXECUTION_ERROR_NUMBER, , sEXECUTION_ERROR_TEXT
   
      If Not _
   oC_VersionControlExport.bExportWorksheetNameData() _
      Then Err.Raise lEXECUTION_ERROR_NUMBER, , sEXECUTION_ERROR_TEXT
   
      If Not _
   oC_VersionControlExport.bExportReferenceData() _
      Then Err.Raise lEXECUTION_ERROR_NUMBER, , sEXECUTION_ERROR_TEXT

Finally:
   On Error Resume Next
   With Application
      .ScreenUpdating = True
      .Calculation = xlCalculationAutomatic
      .EnableEvents = True
      .Calculate
   End With
   Exit Sub

Catch:
   If Not bResumedOnce Then
      bResumedOnce = True: Stop: Resume
   Else
      MsgBox Err.Number & ", " & Err.Description, vbCritical
      Resume Finally
   End If
End Sub


