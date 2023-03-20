Attribute VB_Name = "DEV_f_pM_Testing"
' Purpose: Tests for framework code
' 0.2.0    20.03.2023    gueleh    Initially created
Option Explicit
Option Private Module

Private Const s_m_COMPONENT_NAME As String = "DEV_f_pM_Testing"

Private Sub mTest_f_C_Wks()
   Dim oCWks As New f_C_Wks
   Dim oRng As Range
   Dim oRngEndColumn As Range
   oCWks.Construct DEV_a_wks_TestCanvas
   oCWks.DeleteAllContents
   oCWks.oWks_prop_r.Range("A1:E7").Value = "Test"
   oCWks.oWks_prop_r.Range("B8").Interior.Color = vbWhite
   
   Set oRng = oCWks.oWks_prop_r.Range("B3")
   Set oRngEndColumn = oCWks.oWks_prop_r.Range("D3")
   
   Debug.Print oCWks.oRng_prop_r_CurrentRegionEnhanced(oRng, oRngEndColumn).Address
   
   'Code stops here if test fails
   Debug.Assert oCWks.oRng_prop_r_CurrentRegionEnhanced(oRng, oRngEndColumn).Address = "$B$3:$D$7"
   
   oCWks.SetDataRangeByAnchors oRng, oRngEndColumn
   'Code stops here if test fails
   Debug.Assert oCWks.oRng_prop_rw_Data.Address = "$B$3:$D$7"

End Sub

