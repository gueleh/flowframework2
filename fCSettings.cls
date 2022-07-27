VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "fCSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' -------------------------------------------------------------------------------------------
' CORE, do not change
'============================================================================================
'   NAME:     fCSettings
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
'   0.1.0    20220709    gueleh    Initially created
'--------------------------------------------------------------------------------------------
'   BACKLOG
'   ''''''''''''''''''''
'   none
'============================================================================================
Option Explicit

Private Const s_m_COMPONENT_NAME As String = "fCSettings"

Private Const sm_name_VERSION_NUMBER As String = "fn_smVERSION_NUMBER"
Private Const sm_name_VERSION_DATE_YYMMDD As String = "fn_smVERSION_DATE_YYMMDD"
Private Const sm_name_DEBUG_MODE_IS_ON As String = "afn_bmDebugModeIsOn"

Private omWksFrameworkSettings As Worksheet
Private omWksApplicationSpecificFrameworkSettings As Worksheet
Private smVersionNumber As String
Private smVersionDateYYMMDD As String
Private bmDebugModeIsOn As Boolean
Private bmThisIsATestRun As Boolean

Public Property Get bDebugModeIsOn() As Boolean
   bDebugModeIsOn = bmDebugModeIsOn
End Property

Public Property Get bThisIsATestRun() As Boolean
   bThisIsATestRun = bmThisIsATestRun
End Property

Public Property Let bThisIsATestRun(ByVal bNewValue As Boolean)
   bmThisIsATestRun = bNewValue
End Property

Public Property Get sVersionNumber() As String
   sVersionNumber = smVersionNumber
End Property

Public Property Get sVersionDateYYMMDD() As String
   sVersionDateYYMMDD = smVersionDateYYMMDD
End Property

Private Sub Class_Initialize()
   Set omWksFrameworkSettings = fwksSettings
   Set omWksApplicationSpecificFrameworkSettings = afwksSettings
   smVersionNumber = omWksFrameworkSettings.Names(sm_name_VERSION_NUMBER).RefersToRange.Value2
   smVersionDateYYMMDD = omWksFrameworkSettings.Names(sm_name_VERSION_DATE_YYMMDD).RefersToRange.Value2
   mGetDebugMode
End Sub

Public Sub SetDebugModeIsOnTo(ByVal bNewValue As Boolean)
   bmDebugModeIsOn = bNewValue
   omWksApplicationSpecificFrameworkSettings.Names(sm_name_DEBUG_MODE_IS_ON).RefersToRange.Value2 = bmDebugModeIsOn
End Sub

Private Sub mGetDebugMode()
   bmDebugModeIsOn = omWksApplicationSpecificFrameworkSettings.Names(sm_name_DEBUG_MODE_IS_ON).RefersToRange.Value2
End Sub

