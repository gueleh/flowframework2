Attribute VB_Name = "devfMUserInterface"
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

Private Const s_m_COMPONENT_NAME As String = "devfUserInterface"

' Purpose: write done stamp into active row of devlog
' 0.2.0    20220711    gueleh    Initially created
Public Sub devfMarkLineItemAsDoneInDevLog()
Attribute devfMarkLineItemAsDoneInDevLog.VB_Description = "Marks the line item with Done"
Attribute devfMarkLineItemAsDoneInDevLog.VB_ProcData.VB_Invoke_Func = "m\n14"
   Dim rng As Range
   Dim wks As Worksheet
   Set rng = Selection
   Set wks = rng.Parent
   If wks.Name = devfwksDevLog.Name _
   And rng.Rows.Count = 1 _
   And rng.Row > 2 _
   And wks.Cells(rng.Row, 1).Value2 <> "" Then
      f_g_InitGlobals
      wks.Cells(rng.Row, 4) = oC_f_g_FrameworkSettings.sVersionNumber
      wks.Cells(rng.Row, 5) = oC_f_g_FrameworkSettings.sVersionDateYYMMDD
      wks.Cells(rng.Row, 6) = "Done"
   End If
End Sub
