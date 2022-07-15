Attribute VB_Name = "fpMUtilities"
' -------------------------------------------------------------------------------------------
' CORE, do not change
'============================================================================================
'   NAME:     fpMUtilities
'============================================================================================
'   Purpose:  utilities being part of the core of the template
'   Access:   Public
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

Private Const smCOMPONENT_NAME As String = "fpMUtilities"

' Purpose: Return string for Application.Run for procedures in this workbook
' 0.1.0    20220709    gueleh    Initially created
Public Function sfRunMyProcedure(ByVal sProcedureName As String) As String
  sfRunMyProcedure = "'" & ThisWorkbook.Name & "'!" & sProcedureName
End Function

