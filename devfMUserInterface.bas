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
' 0.10.0    04.08.2022    gueleh    Changed property names to meet new convention
' 0.7.0    02.08.2022    gueleh    Refactored name to match convention
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
Public Sub DEV_f_g_MarkLineItemAsDoneInAfDevLog()
Attribute DEV_f_g_MarkLineItemAsDoneInAfDevLog.VB_Description = "Marks the line item with Done"
Attribute DEV_f_g_MarkLineItemAsDoneInAfDevLog.VB_ProcData.VB_Invoke_Func = "m\n14"
   Dim oRng As Range
   Dim oWks As Worksheet
   Set oRng = Selection
   Set oWks = oRng.Parent
   If oWks.Name = devafwksDevLog.Name _
   And oRng.Rows.Count = 1 _
   And oRng.Row > 2 _
   And oWks.Cells(oRng.Row, 1).Value2 <> "" Then
      f_p_InitGlobals
      oWks.Cells(oRng.Row, 4) = oC_f_p_FrameworkSettings.s_prop_r_VersionNumber
      oWks.Cells(oRng.Row, 5) = oC_f_p_FrameworkSettings.s_prop_r_VersionDateYYMMDD
      oWks.Cells(oRng.Row, 6) = "Done"
   End If
End Sub
