VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "f_C_Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' -------------------------------------------------------------------------------------------
' CORE, do not change
'============================================================================================
'   NAME:     f_C_Settings
'============================================================================================
'   Purpose:  handling of the framework settings
'   Access:   Public
'   Type:     Class Module
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
' 0.8.0    03.08.2022    gueleh    Added SetNamesVisibleTo
' 0.7.0    02.08.2022    gueleh    Refactored names
'   0.1.0    20220709    gueleh    Initially created
'--------------------------------------------------------------------------------------------
'   BACKLOG
'   ''''''''''''''''''''
'   none
'============================================================================================
Option Explicit

Private Const s_m_COMPONENT_NAME As String = "f_C_Settings"

Private Const s_m_named_cell_s_m_VERSION_NUMBER As String = "f_named_cell_s_m_VERSION_NUMBER"
Private Const s_m_named_cell_s_m_VERSION_DATE_YYMMDD As String = "f_named_cell_s_m_VERSION_DATE_YYMMDD"
Private Const s_m_named_cell_b_m_DEBUG_MODE_IS_ON As String = "af_named_cell_b_m_DebugModeIsOn"
Private Const s_m_named_cell_b_m_DEV_MODE_IS_ON As String = "af_named_cell_b_m_DevelopmentModeIsOn"
Private Const s_m_named_cell_b_m_MAINTENANCE_MODE_IS_ON As String = "af_named_cell_b_m_MaintenanceModeIsOn"

Private oWks_m_FrameworkSettings As Worksheet
Private oWks_m_ApplicationSpecificFrameworkSettings As Worksheet
Private s_m_VersionNumber As String
Private s_m_VersionDateYYMMDD As String
Private b_m_DebugModeIsOn As Boolean
Private b_m_ThisIsATestRun As Boolean
Private b_m_DevelopmentModeIsOn As Boolean
Private b_m_MaintenanceModeIsOn As Boolean

Public Property Get b_prop_r_DebugModeIsOn() As Boolean
   b_prop_r_DebugModeIsOn = b_m_DebugModeIsOn
End Property

Public Property Get b_prop_r_DevelopmentModeIsOn() As Boolean
   b_prop_r_DevelopmentModeIsOn = b_m_DevelopmentModeIsOn
End Property

Public Property Get b_prop_r_MaintenanceModeIsOn() As Boolean
   b_prop_r_MaintenanceModeIsOn = b_m_MaintenanceModeIsOn
End Property

Public Property Get b_prop_rw_ThisIsATestRun() As Boolean
   b_prop_rw_ThisIsATestRun = b_m_ThisIsATestRun
End Property

Public Property Let b_prop_rw_ThisIsATestRun(ByVal bNewValue As Boolean)
   b_m_ThisIsATestRun = bNewValue
End Property

Public Property Get s_prop_r_VersionNumber() As String
   s_prop_r_VersionNumber = s_m_VersionNumber
End Property

Public Property Get s_prop_r_VersionDateYYMMDD() As String
   s_prop_r_VersionDateYYMMDD = s_m_VersionDateYYMMDD
End Property

Private Sub Class_Initialize()
   Set oWks_m_FrameworkSettings = f_wks_Settings
   Set oWks_m_ApplicationSpecificFrameworkSettings = af_wks_Settings
   s_m_VersionNumber = oWks_m_FrameworkSettings.Names(s_m_named_cell_s_m_VERSION_NUMBER).RefersToRange.Value2
   s_m_VersionDateYYMMDD = oWks_m_FrameworkSettings.Names(s_m_named_cell_s_m_VERSION_DATE_YYMMDD).RefersToRange.Value2
   mGetDebugMode
   mGetDevelopmentMode
   mGetMaintenanceMode
End Sub

Public Sub SetDebugModeIsOnTo(ByVal bNewValue As Boolean)
   b_m_DebugModeIsOn = bNewValue
   oWks_m_ApplicationSpecificFrameworkSettings.Names(s_m_named_cell_b_m_DEBUG_MODE_IS_ON).RefersToRange.Value2 = b_m_DebugModeIsOn
End Sub

Public Sub SetDevelopmentModeIsOnTo(ByVal bNewValue As Boolean)
   b_m_DevelopmentModeIsOn = bNewValue
   oWks_m_ApplicationSpecificFrameworkSettings.Names(s_m_named_cell_b_m_DEV_MODE_IS_ON).RefersToRange.Value2 = b_m_DevelopmentModeIsOn
End Sub

Public Sub SetMaintenanceModeIsOnTo(ByVal bNewValue As Boolean)
   b_m_MaintenanceModeIsOn = bNewValue
   oWks_m_ApplicationSpecificFrameworkSettings.Names(s_m_named_cell_b_m_MAINTENANCE_MODE_IS_ON).RefersToRange.Value2 = b_m_MaintenanceModeIsOn
End Sub

' Purpose: setting the visibility of the Names used in this class
' 0.8.0    03.08.2022    gueleh    Initially created
Public Sub SetNamesVisibleTo(ByVal bVisible As Boolean)
   oWks_m_ApplicationSpecificFrameworkSettings.Names(s_m_named_cell_b_m_DEBUG_MODE_IS_ON).Visible = bVisible
   oWks_m_ApplicationSpecificFrameworkSettings.Names(s_m_named_cell_b_m_DEV_MODE_IS_ON).Visible = bVisible
   oWks_m_FrameworkSettings.Names(s_m_named_cell_s_m_VERSION_NUMBER).Visible = bVisible
   oWks_m_FrameworkSettings.Names(s_m_named_cell_s_m_VERSION_DATE_YYMMDD).Visible = bVisible
   'The visibility of Names of ListObjects depends on the visibility of the worksheet they are in
End Sub

Private Sub mGetDebugMode()
   b_m_DebugModeIsOn = oWks_m_ApplicationSpecificFrameworkSettings.Names(s_m_named_cell_b_m_DEBUG_MODE_IS_ON).RefersToRange.Value2
End Sub

Private Sub mGetDevelopmentMode()
   b_m_DevelopmentModeIsOn = oWks_m_ApplicationSpecificFrameworkSettings.Names(s_m_named_cell_b_m_DEV_MODE_IS_ON).RefersToRange.Value2
End Sub

Private Sub mGetMaintenanceMode()
   b_m_MaintenanceModeIsOn = oWks_m_ApplicationSpecificFrameworkSettings.Names(s_m_named_cell_b_m_MAINTENANCE_MODE_IS_ON).RefersToRange.Value2
End Sub

