Attribute VB_Name = "DEV_a_pM_SandBox"
' Purpose: Sandbox for dev experiments
' 0.1.0    17.03.2023    gueleh    Initially created
Option Explicit
Option Private Module

Private Const s_m_COMPONENT_NAME As String = "DEV_a_pM_SandBox"

Private Sub mTest_f_C_Wks()
   Dim oCWks As New f_C_Wks
   Dim oRng As Range
   Dim oRngEndColumn As Range
   oCWks.Construct DEV_a_wks_TestCanvas
   oCWks.DeleteAllContents
   oCWks.oWks_prop_r.Range("A1:E7").Value = "Test"
   Set oRng = oCWks.oWks_prop_r.Range("B3")
   Set oRngEndColumn = oCWks.oWks_prop_r.Range("D3")
   Debug.Print oCWks.oRng_prop_r_CurrentRegionEnhanced(oRng, oRngEndColumn).Address
End Sub
