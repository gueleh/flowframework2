Attribute VB_Name = "DEV_f_pM_Test_f_C_Wks"
' CORE-DEV - do not change, optionally remove when deploying app
'============================================================================================
'   NAME:     DEV_f_pM_Test_f_C_Wks
'============================================================================================
'   Purpose:  directly accessible dev helpers
'   Access:   Private
'   Type:     Module
'   Author:   Günther Lehner
'   Contact:  guenther.lehner@protonmail.com
'   GitHubID: gueleh
'   Required:
'   Usage:
'  + can be used without the framework, except for a test canvas sheet with the codename DEV_a_wks_TestCanvas
'  + based on Debug.Assert, which stops the code if the test fails
'--------------------------------------------------------------------------------------------
'   VERSION HISTORY
'   Version    Date    Developer    Changes
'   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'     0.2.0    20.03.2023    gueleh    Initially created
'--------------------------------------------------------------------------------------------
'   BACKLOG
'   ''''''''''''''''''''
'   none
'============================================================================================
Option Explicit
Option Private Module

Private Const s_m_COMPONENT_NAME As String = "DEV_f_pM_Test_f_C_Wks"

Private b_m_ChangeEventInvoked As Boolean

Private Sub mTest_f_C_Wks()
   Dim oCWks As New f_C_Wks
   Dim oRng As Range
   Dim oRngEndColumn As Range
   Dim oRngCell As Range
   
   b_m_ChangeEventInvoked = False
   
   oCWks.Construct DEV_a_wks_TestCanvas
   oCWks.DeleteAllContents
   For Each oRngCell In oCWks.oWks_prop_r.Range("A1:E7")
      oRngCell.Value2 = "Test-" & oRngCell.Address
   Next oRngCell
   oCWks.oWks_prop_r.Range("B8").Interior.Color = vbWhite
   
   Set oRng = oCWks.oWks_prop_r.Range("B3")
   Set oRngEndColumn = oCWks.oWks_prop_r.Range("D3")
   
   'Test: setting current region based on anchor cells (i.e. omitting excess data due to not empty adjacent cells)
   'Code stops here if test fails
   Debug.Assert oCWks.oRng_prop_r_CurrentRegionEnhanced(oRng, oRngEndColumn).Address = "$B$3:$D$7"
   
   'Test: setting the data range in the class, without the first row of it being the header row
   oCWks.SetDataRangeByAnchors oRng, oRngEndColumn
   'Code stops here if test fails
   Debug.Assert oCWks.oRng_prop_r_Data.Address = "$B$3:$D$7"

   'Test: setting the header row dict with pointers to the columns by names
   '  Note: the header row is above the data range,
   '  there's even a row inbetween: header in row 1, data beginning in row 3
   oCWks.CreateHeaderDictionary 1
   'Code stops here if test fails
   Debug.Assert oCWks.l_prop_r_ColumnNumberByHeaderName("Test-$D$1") = 4
   
   'Test: sanitize the used range, which ends in row 8 before the sanitation
   oCWks.SanitizeUsedRange
   'Code stops here if test fails
   Debug.Assert oCWks.oWks_prop_r.UsedRange.Address = "$A$1:$E$7"
   
   oCWks.SetDataRangeByAnchors oRng, oRngEndColumn, True, True
   'Code stops here if test fails
   Debug.Assert oCWks.l_prop_r_ColumnNumberByHeaderName("Test-$C$3") = 3
   
   oCWks.s_prop_rw_NameOfSubToRunOnWksChange = "mOnChangeTest"
   'should not invoke the on change sub
   oCWks.oWks_prop_r.Range("$L$1").Value2 = "Change!"
   Debug.Assert b_m_ChangeEventInvoked = False
   
   'should invoke the on change sub
   oCWks.b_prop_rw_WksChangeEventIsActive = True
   oCWks.oWks_prop_r.Range("$L$1").Value2 = "Change!"
   Debug.Assert b_m_ChangeEventInvoked = True
   
End Sub

Private Sub mOnChangeTest(ByRef oRng_arg_Target As Range)
   b_m_ChangeEventInvoked = True
   Debug.Assert oRng_arg_Target.Address = "$L$1"
End Sub
