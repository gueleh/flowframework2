VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "a_C_AppSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Purpose: the app settings in a trivial class with public fields
Option Explicit

Private Const s_m_NAME_s_name_wkb_AppVersionNumber As String = "s_name_wkb_AppVersionNumber"
Private Const s_m_NAME_dte_name_wkb_AppVersionDate As String = "dte_name_wkb_AppVersionDate"

Public sAppVersionNumber As String
Public dteAppVersionDate As String

Private Sub Class_Initialize()
   sAppVersionNumber = v_f_p_ValueFromWorkbookName(s_m_NAME_s_name_wkb_AppVersionNumber)
   dteAppVersionDate = v_f_p_ValueFromWorkbookName(s_m_NAME_dte_name_wkb_AppVersionDate)
End Sub
