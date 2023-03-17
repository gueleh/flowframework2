Attribute VB_Name = "af_pM_Globals"
' -------------------------------------------------------------------------------------------
' APP-SPECIFIC CORE, your content has to be migrated manually if the template if updated
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
' 0.1.0    17.03.2023    gueleh    Initially created
'--------------------------------------------------------------------------------------------
'   BACKLOG
'   ''''''''''''''''''''
'   none
'============================================================================================
Option Explicit
Option Private Module

Private Const s_m_COMPONENT_NAME As String = "af_pM_Globals"

Public Const s_p_PASSWORD As String = "OxesAndFrogsDanceToTrance"

' When debug mode is on, the code stops once in case of an error so _
' that you can see in which row the error did occur
Public Const b_p_DEBUG_MODE_IS_ON As Boolean = True

' When sheet protection is enabled, all sheets will be protected when processing ends
Public Const b_p_ENABLE_SHEET_PROTECTION As Boolean = False

' Purpose: speeds up Excel and deactivates certain functionalities,
'           unprotects all worksheets
'           and sets the cancel to error handling if the Const for the debug mode is set to true
' -------------------
' ADAPT TO YOUR NEEDS
' -------------------
' 0.1.0    17.03.2023    gueleh    Initially created
Public Sub af_p_StartProcessing()
   Dim oWks As Worksheet
   
   With Application
      If b_p_DEBUG_MODE_IS_ON Then
         .EnableCancelKey = xlErrorHandler
      Else
         .EnableCancelKey = xlDisabled
      End If
      .ScreenUpdating = False
      .Calculation = xlCalculationManual
      .EnableEvents = False
   End With
   
   For Each oWks In ThisWorkbook.Worksheets
      If oWks.ProtectContents Then oWks.Unprotect s_p_PASSWORD
   Next oWks
End Sub

' Purpose: activates all functionalities again and protects all worksheets if
'           if the Const above is set to true
' -------------------
' ADAPT TO YOUR NEEDS
' -------------------
' 0.1.0    17.03.2023    gueleh    Initially created
Public Sub af_p_EndProcessing()
   Dim oWks As Worksheet
   
   If b_p_ENABLE_SHEET_PROTECTION Then
      For Each oWks In ThisWorkbook.Worksheets
         oWks.Protect _
            DrawingObjects:=True, _
            Contents:=True, _
            Scenarios:=True, _
            AllowFormattingColumns:=True, _
            AllowFormattingRows:=True, _
            AllowSorting:=True, _
            AllowFiltering:=True, _
            AllowUsingPivotTables:=True
      Next oWks
   End If
   
   With Application
      .EnableCancelKey = xlInterrupt
      .ScreenUpdating = True
      .Calculation = xlCalculationAutomatic
      .EnableEvents = True
      .Calculate
   End With

End Sub

