VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "fclsSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' -------------------------------------------------------------------------------------------
' CORE, do not change
'============================================================================================
'   NAME:     fclsSettings
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

Private Const msCOMPONENT_NAME As String = "fclsSettings"

Private Const ms_name_VERSION_NUMBER As String = "fn_msVERSION_NUMBER"
Private Const ms_name_VERSION_DATE_YYMMDD As String = "fn_msVERSION_DATE_YYMMDD"
Private Const ms_name_DEBUG_MODE_IS_ON As String = "afn_mbDebugModeIsOn"

Private mwksFrameworkSettings As Worksheet
Private mwksApplicationSpecificFrameworkSettings As Worksheet
Private msVersionNumber As String
Private msVersionDateYYMMDD As String
Private mbDebugModeIsOn As Boolean
Private mbThisIsATestRun As Boolean

Public Property Get bDebugModeIsOn() As Boolean
   bDebugModeIsOn = mbDebugModeIsOn
End Property

Public Property Get bThisIsATestRun() As Boolean
   bThisIsATestRun = mbThisIsATestRun
End Property

Public Property Let bThisIsATestRun(ByVal bNewValue As Boolean)
   mbThisIsATestRun = bNewValue
End Property

Public Property Get sVersionNumber() As String
   sVersionNumber = msVersionNumber
End Property

Public Property Get sVersionDateYYMMDD() As String
   sVersionDateYYMMDD = msVersionDateYYMMDD
End Property

Private Sub Class_Initialize()
   Set mwksFrameworkSettings = fwksSettings
   Set mwksApplicationSpecificFrameworkSettings = afwksSettings
   msVersionNumber = mwksFrameworkSettings.Names(ms_name_VERSION_NUMBER).RefersToRange.Value2
   msVersionDateYYMMDD = mwksFrameworkSettings.Names(ms_name_VERSION_DATE_YYMMDD).RefersToRange.Value2
   mGetDebugMode
End Sub

Public Sub SetDebugModeIsOnTo(ByVal bNewValue As Boolean)
   mbDebugModeIsOn = bNewValue
   mwksApplicationSpecificFrameworkSettings.Names(ms_name_DEBUG_MODE_IS_ON).RefersToRange.Value2 = mbDebugModeIsOn
End Sub

Private Sub mGetDebugMode()
   mbDebugModeIsOn = mwksApplicationSpecificFrameworkSettings.Names(ms_name_DEBUG_MODE_IS_ON).RefersToRange.Value2
End Sub

