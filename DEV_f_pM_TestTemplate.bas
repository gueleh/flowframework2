Attribute VB_Name = "DEV_f_pM_TestTemplate"

' CORE-DEV - do not change, optionally remove when deploying app
'============================================================================================
'   NAME:     DEV_f_pM_TestTemplate
'============================================================================================
Option Explicit
Option Private Module

Private Const s_m_COMPONENT_NAME As String = "DEV_f_pM_TestTemplate"

Public Sub mTest_MeinFeature()
   Dim oC_Test As DEV_f_C_UnitTest
   Dim sErgebnis As String
   Set oC_Test = oC_DEV_f_p_CreateTest("Beschreibung", s_m_COMPONENT_NAME, "mTest_MeinFeature")
   
   ' Arrange / Act / Assert
   oC_Test.oC_prop_r_Assert.AssertEqual "erwartet", sErgebnis, "Erklärung"
   
   DEV_f_p_CompleteTest oC_Test
End Sub

