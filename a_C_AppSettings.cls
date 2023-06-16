VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "a_C_AppSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Belongs to APP - will not be updated when updating the framework
' Purpose: the app settings in a trivial class with public fields
Option Explicit

Private Const s_m_NAME_a_named_cell_s_g_AppVersionNumber As String = "a_named_cell_s_g_AppVersionNumber"
Private Const s_m_NAME_a_named_cell_dte_g_AppVersionDate As String = "a_named_cell_dte_g_AppVersionDate"

Public sAppVersionNumber As String
Public dteAppVersionDate As String

Private Sub Class_Initialize()
   sAppVersionNumber = v_f_p_ValueFromWorkbookName(s_m_NAME_a_named_cell_s_g_AppVersionNumber)
   dteAppVersionDate = v_f_p_ValueFromWorkbookName(s_m_NAME_a_named_cell_dte_g_AppVersionDate)
End Sub
