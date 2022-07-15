Attribute VB_Name = "devfUserInterface"
' -------------------------------------------------------------------------------------------
' CORE-DEV, do not change
'============================================================================================
'   NAME:     devfUserInterface
'============================================================================================
'   Purpose:  directly accessible dev helpers
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
'   0.2.0    20220711    gueleh    Initially created
'--------------------------------------------------------------------------------------------
'   BACKLOG
'   ''''''''''''''''''''
'   none
'============================================================================================
Option Explicit

Private Const msCOMPONENT_NAME As String = "devfUserInterface"

' Purpose: write done stamp into active row of devlog
' 0.2.0    20220711    gueleh    Initially created
Public Sub MarkLineItemAsDoneInDevLog()
Attribute MarkLineItemAsDoneInDevLog.VB_Description = "Marks the line item with Done"
Attribute MarkLineItemAsDoneInDevLog.VB_ProcData.VB_Invoke_Func = "m\n14"
   Dim rng As Range
   Dim wks As Worksheet
   Set rng = Selection
   Set wks = rng.Parent
   If wks.Name = devfwksDevLog.Name _
   And rng.Rows.Count = 1 _
   And rng.Row > 2 _
   And wks.Cells(rng.Row, 1).Value2 <> "" Then
      fInitGlobals
      wks.Cells(rng.Row, 4) = fgclsFrameworkSettings.sVersionNumber
      wks.Cells(rng.Row, 5) = fgclsFrameworkSettings.sVersionDateYYMMDD
      wks.Cells(rng.Row, 6) = "Done"
   End If
End Sub
