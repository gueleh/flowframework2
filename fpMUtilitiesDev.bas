Attribute VB_Name = "fpMUtilitiesDev"
' -------------------------------------------------------------------------------------------
' CORE, do not change
'============================================================================================
'   NAME:     fpMUtilitiesDev
'============================================================================================
'   Purpose:  utilities for development work which have to be available also in production
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
' 0.10.0    04.08.2022    gueleh    Changed property names to meet new convention
' 0.9.0    03.08.2022    gueleh    Added b_f_p_SetDevelopmentModeTo
'   0.8.0    03.08.2022    gueleh    Initially created
'--------------------------------------------------------------------------------------------
'   BACKLOG
'   ''''''''''''''''''''
'   none
'============================================================================================
Option Explicit
Option Private Module

Private Const s_m_COMPONENT_NAME As String = "fpMUtilitiesDev"

Public Sub f_SetTechnicalNamesVisibleToFalse()
   Dim oC As New fCSettings
   oC.SetNamesVisibleTo False
End Sub

Public Sub f_SetTechnicalNamesVisibleToTrue()
   Dim oC As New fCSettings
   oC.SetNamesVisibleTo True
End Sub

' Purpose: setting the development mode to the provided value, see also caller doc in fpMEntryLevel
' 0.9.0    03.08.2022    gueleh    Initially created
Public Function b_f_p_SetDevelopmentModeTo(ByVal bDevModeIsOn As Boolean) As Boolean

   Dim oC_Me As New fCCallParams
   oC_Me.s_prop_rw_ComponentName = s_m_COMPONENT_NAME
   If oC_f_p_FrameworkSettings.b_prop_rw_ThisIsATestRun Then f_p_RegisterUnitTest oC_Me

   With oC_Me
      .s_prop_rw_ProcedureName = "b_f_p_SetDevelopmentModeTo" 'Name of the function
      .b_prop_rw_SilentError = True 'False will display a message box - you should only do this on entry level
      .s_prop_rw_ErrorMessage = "Setting the development mode failed." 'A message that properly informs the user and the devs (silent errors will be logged nonetheless)
      .SetCallArgs "Set to: " & bDevModeIsOn 'If the sub takes args put the here like ("sExample:=" & sExample, "lExample:=" & lExample)
   End With

Try:
   On Error GoTo Catch
   Dim saTechWksIdentifiers() As String
   Dim oWks As Worksheet
   Dim l As Long
   Dim eVisibility As XlSheetVisibility
   
   If bDevModeIsOn Then
      eVisibility = xlSheetVisible
   Else
      eVisibility = xlSheetVeryHidden
   End If
   saTechWksIdentifiers = Split(s_f_p_split_seed_TECH_WKS_IDENTIFIERS, s_f_p_SPLIT_SEED_SEPARATOR)
   
   For Each oWks In ThisWorkbook.Worksheets
      For l = LBound(saTechWksIdentifiers) To UBound(saTechWksIdentifiers)
         If Left$(oWks.CodeName, Len(saTechWksIdentifiers(l))) = saTechWksIdentifiers(l) Then
            oWks.Visible = eVisibility
         End If
      Next l
   Next oWks
   
   oC_f_p_FrameworkSettings.SetNamesVisibleTo bDevModeIsOn
   oC_f_p_FrameworkSettings.SetDevelopmentModeIsOnTo bDevModeIsOn
   If Not bDevModeIsOn Then oC_f_p_FrameworkSettings.SetDebugModeIsOnTo bDevModeIsOn
   
Finally:
   On Error Resume Next
   If oC_Me.oC_prop_r_Error Is Nothing Then b_f_p_SetDevelopmentModeTo = True 'reports execution as successful to caller
   Exit Function
   
Catch:
   If oC_Me.oC_prop_r_Error Is Nothing _
   Then f_p_RegisterError oC_Me, Err.Number, Err.Description
   If oC_f_p_FrameworkSettings.b_prop_rw_ThisIsATestRun Then f_p_RegisterExecutionError oC_Me
   If oC_f_p_FrameworkSettings.b_prop_r_DebugModeIsOn And Not oC_Me.b_prop_rw_ResumedOnce Then
      oC_Me.b_prop_rw_ResumedOnce = True: Stop: Resume
   Else
      f_p_HandleError oC_Me: Resume Finally
   End If

End Function


