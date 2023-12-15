VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "af_C_AppModes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' -------------------------------------------------------------------------------------------
' customizable part of the framework
'============================================================================================
'   NAME:     af_C_AppModes
'============================================================================================
'   Purpose:  handles production, maintenance and dev mode switching
'   Access:   Public
'   Type:     Class Module
'   Author:   Günther Lehner
'   Contact:  guleh@pm.me
'   GitHubID:
'   Required:
'   Usage:
'--------------------------------------------------------------------------------------------
'   VERSION HISTORY
'   Version    Date    Developer    Changes
'   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   1.10.0    23.11.2023        Initially created
'--------------------------------------------------------------------------------------------
'   BACKLOG
'   ''''''''''''''''''''
'   none
'============================================================================================
Option Explicit

Private Const s_m_COMPONENT_NAME As String = "af_C_AppModes"

Private oCol_m_MaintenanceSheets As Collection
Private oCol_m_DevelopmentSheets As Collection

'>>>>>>>>>>>>>>>> CUSTOMIZABLE PART OF THE CLASS
Private Const sPASSWORD_DEV_MODE As String = "your-password-here"
Private Const sPASSWORD_MAINTENANCE_MODE As String = "your-password-here"

' Purpose: builds collection with sheets which are supposed to be visible
'     only in maintenance and dev mode
' 1.10.0    23.11.2023    gueleh    Initially created
Private Function oCol_m_GetMaintenanceSheets() As Collection
   Dim oCol As New Collection
Try:
   On Error GoTo Catch
   With oCol
      .Add a_wks_Administration
      
'>>>>>>> Your development sheets here
' all the sheets which are supposed to be visible only in maintenance and dev mode
      
'<<<<<<<
   End With
   Set oCol_m_GetMaintenanceSheets = oCol
   Exit Function
Catch:
   Set oCol_m_GetMaintenanceSheets = New Collection
End Function

' Purpose: builds collection with sheets which are supposed to be very hidden when not in dev mode
' 1.10.0    23.11.2023    gueleh    Initially created
Private Function oCol_m_GetDevelopmentSheets() As Collection
   Dim oCol As New Collection
Try:
   On Error GoTo Catch
   With oCol
      .Add a_wks_Settings
      .Add a_wks_VersionControlRanges
      .Add af_wks_ErrorLog
      .Add af_wks_Settings
      .Add DEV_a_wks_TestCanvas
      .Add DEV_f_wks_TestCanvas
      .Add f_wks_Settings
      
'>>>>>>> Your development sheets here
' all the sheets which are supposed to be very hidden when app is not in dev mode
      
'<<<<<<<
   End With
   Set oCol_m_GetDevelopmentSheets = oCol
   Exit Function
Catch:
   Set oCol_m_GetDevelopmentSheets = New Collection
End Function

'<<<<<<<<<<<<<<<< END OF CUSTOMIZABLE PART OF THE CLASS


Private Sub Class_Initialize()
   Set oCol_m_DevelopmentSheets = oCol_m_GetDevelopmentSheets
   Set oCol_m_MaintenanceSheets = oCol_m_GetMaintenanceSheets
End Sub

Public Function bPasswordDevModeCorrect(ByVal sPassword As String) As Boolean
   If sPassword = sPASSWORD_DEV_MODE Then bPasswordDevModeCorrect = True
End Function

Public Function bPasswordMaintenanceModeCorrect(ByVal sPassword As String) As Boolean
   If sPassword = sPASSWORD_MAINTENANCE_MODE Then bPasswordMaintenanceModeCorrect = True
End Function

Public Function bSetDevelopmentModeTo(ByVal bDevModeIsOn As Boolean) As Boolean

   Dim oC_Me As New f_C_CallParams
   oC_Me.s_prop_rw_ComponentName = s_m_COMPONENT_NAME
   If oC_f_p_FrameworkSettings.b_prop_rw_ThisIsATestRun Then f_p_RegisterUnitTest oC_Me

   With oC_Me
      .s_prop_rw_ProcedureName = "bSetDevelopmentModeTo" 'Name of the function
      .b_prop_rw_SilentError = True 'False will display a message box - you should only do this on entry level
      .s_prop_rw_ErrorMessage = "Setting the development mode failed." 'A message that properly informs the user and the devs (silent errors will be logged nonetheless)
      .SetCallArgs "Set to: " & bDevModeIsOn 'If the sub takes args put the here like ("sExample:=" & sExample, "lExample:=" & lExample)
   End With

Try:
   On Error GoTo Catch
   Dim oWks As Worksheet
   Dim eVisibility As XlSheetVisibility
   
   If bDevModeIsOn Then
      eVisibility = xlSheetVisible
   Else
      eVisibility = xlSheetVeryHidden
   End If
   
   For Each oWks In oCol_m_DevelopmentSheets
      oWks.Visible = eVisibility
   Next oWks
   
   oC_f_p_FrameworkSettings.SetNamesVisibleTo bDevModeIsOn
   oC_f_p_FrameworkSettings.SetDevelopmentModeIsOnTo bDevModeIsOn
   If Not bDevModeIsOn Then oC_f_p_FrameworkSettings.SetDebugModeIsOnTo bDevModeIsOn
   
Finally:
   On Error Resume Next
   If oC_Me.oC_prop_r_Error Is Nothing Then bSetDevelopmentModeTo = True 'reports execution as successful to caller
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

Public Function bSetMaintenanceModeTo(ByVal bModeIsOn As Boolean) As Boolean

   Dim oC_Me As New f_C_CallParams
   oC_Me.s_prop_rw_ComponentName = s_m_COMPONENT_NAME
   If oC_f_p_FrameworkSettings.b_prop_rw_ThisIsATestRun Then f_p_RegisterUnitTest oC_Me

   With oC_Me
      .s_prop_rw_ProcedureName = "bSetMaintenanceModeTo" 'Name of the function
      .b_prop_rw_SilentError = True 'False will display a message box - you should only do this on entry level
      .s_prop_rw_ErrorMessage = "Setting the maintenance mode failed." 'A message that properly informs the user and the devs (silent errors will be logged nonetheless)
      .SetCallArgs "Set to: " & bModeIsOn 'If the sub takes args put the here like ("sExample:=" & sExample, "lExample:=" & lExample)
   End With

Try:
   On Error GoTo Catch
   Dim oWks As Worksheet
   Dim eVisibility As XlSheetVisibility
   
   If bModeIsOn Then
      eVisibility = xlSheetVisible
   Else
      eVisibility = xlSheetVeryHidden
   End If
   
   For Each oWks In oCol_m_MaintenanceSheets
      oWks.Visible = eVisibility
   Next oWks
   
   oC_f_p_FrameworkSettings.SetMaintenanceModeIsOnTo bModeIsOn
   
Finally:
   On Error Resume Next
   If oC_Me.oC_prop_r_Error Is Nothing Then bSetMaintenanceModeTo = True 'reports execution as successful to caller
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

