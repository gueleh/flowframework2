Attribute VB_Name = "DEV_f_pM_Test_TemplRenderer"
' -------------------------------------------------------------------------------------------
' DEV, remove from production version
'============================================================================================
'   NAME:     DEV_f_pM_Test_TemplRenderer
'============================================================================================
'   Purpose:  demo/test Template Renderer Lite and Template Renderer
'   Access:   Private
'   Type:     Modul
'   Author:   Guenther Lehner
'   Contact:  guleh@pm.me
'   GitHubID: gueleh
'   Required:
'   Usage:
'--------------------------------------------------------------------------------------------
'   VERSION HISTORY
'   Version    Date    Developer    Changes
'   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'--------------------------------------------------------------------------------------------
'   BACKLOG
'   ''''''''''''''''''''
'   none
'============================================================================================
Option Explicit
Option Private Module

Private Const s_m_COMPONENT_NAME As String = "DEV_f_pM_Test_TemplRenderer"

Public Sub DEV_f_p_TRlite_RenderInvoiceExample_Legacy()
    ' 1) Styles aus _meta sicherstellen
    f_p_EnsureStylesFromMeta af_wks_Styles.name
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    On Error GoTo CleanFail
    
    ' 2) Template parsen
    Dim tpl As u_f_TRlite_BlockSpec
    tpl = u_f_p_ParseTemplate(DEV_f_wks_TemplateLite, "blk_Invoice", "rep_Items")
    
    ' 3) Demo-Daten (Dictionarys/Collection) – später durch echtes Binding ersetzen
    Dim header As Object, items As Object, totals As Object
    Set header = CreateObject("Scripting.Dictionary")
    Set totals = CreateObject("Scripting.Dictionary")
    
    header("Invoice.Number") = "INV-2025-091"
    header("Invoice.Date") = Format(Date, "yyyy-mm-dd")
    header("Customer.Name") = "Acme GmbH"
    header("Customer.City") = "Berlin"
    header("Customer.Country") = "DE"
    
    Dim arrItems As Collection: Set arrItems = New Collection
    AddItem arrItems, "Beratung Tagessatz", 2, 1250#
    AddItem arrItems, "Konzept-Workshop", 1, 2200#
    AddItem arrItems, "Dokumentation", 3, 400#
    
    totals("Totals.Sum") = SumItems(arrItems)
    
    ' 4) Output schreiben
    Dim wsOut As Worksheet
    Set wsOut = DEV_f_wks_TestCanvas
    wsOut.Cells.Delete
    
    WriteBlock wsOut, tpl, header, arrItems, totals
    
    ' 5) Autofit
    wsOut.Columns("A:D").EntireColumn.AutoFit
    
CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Exit Sub
CleanFail:
    MsgBox "Render failed: " & Err.Description, vbCritical
    Resume CleanUp
End Sub

Public Sub DEV_f_p_TR_RenderInvoiceExample()
'Fixed, don't change
   Dim oC_Me As New f_C_CallParams: oC_Me.s_prop_rw_ComponentName = s_m_COMPONENT_NAME
'>>>>>>> Your custom settings here
   f_p_StartProcessing e_f_p_ProcessingMode_AutoCalcOffOnSceenUpdatingOffOn
   With oC_Me
      .s_prop_rw_ProcedureName = "DEV_f_p_TR_RenderInvoiceExample" 'Name of the sub
      .b_prop_rw_SilentError = False 'False will display a message box - you should only do this on entry level
      .s_prop_rw_ErrorMessage = "Test of Template Renderer failed." 'A message that properly informs the user and the devs (silent errors will be logged nonetheless)
      .SetCallArgs "No args" 'If the sub takes args put the here like ("sExample:=" & sExample, "lExample:=" & lExample)
   End With
'Fixed, don't change
   If oC_f_p_FrameworkSettings.b_prop_rw_ThisIsATestRun Then f_p_RegisterUnitTest oC_Me
Try: On Error GoTo Catch


'>>>>>>> Your code here
    
   f_p_EnsureStylesFromMeta af_wks_Styles.name
   
   
   Dim wsTpl As Worksheet, wsOut As Worksheet
   Dim uaBlockSpecs() As u_f_BlockSpecRenderer
   Dim data As Object
   Dim lNextStartRow As Long
   
   Set wsTpl = DEV_f_wks_Template
   Set wsOut = DEV_f_wks_TestCanvas
   wsOut.Cells.Delete
      
      If Not _
   b_f_p_GetAllParsedBlockSpecs(uaBlockSpecs, wsTpl) _
      Then Err.Raise _
         e_f_p_HandledError_ExecutionOfLowerLevelFunction, , _
         s_f_p_HandledErrorDescription(e_f_p_HandledError_ExecutionOfLowerLevelFunction)
    
    
   Set data = BuildDemoDataForRenderer2()
   
   lNextStartRow = 1
   
      If Not _
   b_f_p_RenderBlocks(wsOut, uaBlockSpecs, data, lNextStartRow, 1, lNextStartRow) _
      Then Err.Raise _
         e_f_p_HandledError_ExecutionOfLowerLevelFunction, , _
         s_f_p_HandledErrorDescription(e_f_p_HandledError_ExecutionOfLowerLevelFunction)
   
      If Not _
   b_f_p_RenderBlocks(wsOut, uaBlockSpecs, data, (lNextStartRow + 1), 1) _
      Then Err.Raise _
         e_f_p_HandledError_ExecutionOfLowerLevelFunction, , _
         s_f_p_HandledErrorDescription(e_f_p_HandledError_ExecutionOfLowerLevelFunction)
   
   wsOut.Columns.AutoFit
       
   
'End of your code <<<<<<<
   

'Fixed, don't change
Finally: On Error Resume Next


'>>>>>>> Your code here



'End of your code <<<<<<<


'>>>>>>> Your custom settings here
   f_p_EndProcessing e_f_p_ProcessingMode_AutoCalcOffOnSceenUpdatingOffOn
'Fixed, don't change
   Exit Sub
HandleError: af_pM_ErrorHandling.af_p_Hook_ErrorHandling_EntryLevel


'>>>>>>> Your code here



'End of your code <<<<<<<
   
   
'Fixed, don't change
   Resume Finally
Catch:
   If oC_Me.oC_prop_r_Error Is Nothing Then f_p_RegisterError oC_Me, Err.Number, Err.Description
   If oC_f_p_FrameworkSettings.b_prop_rw_ThisIsATestRun Then f_p_RegisterExecutionError oC_Me
   If oC_f_p_FrameworkSettings.b_prop_r_DebugModeIsOn And Not oC_Me.b_prop_rw_ResumedOnce Then
      oC_Me.b_prop_rw_ResumedOnce = True: Stop: Resume
   Else
      f_p_HandleError oC_Me: GoTo HandleError
   End If
End Sub


Private Function BuildDemoDataForRenderer2() As Object
    Dim root As Object: Set root = CreateObject("Scripting.Dictionary")
    
    ' Block: Panel (links fix, rechts rep)
    Dim ctxPanel As Object: Set ctxPanel = oDict_f_p_BuildEmptyContext()
    ctxPanel("header")("Customer.Name") = "Acme GmbH"
    ctxPanel("header")("Customer.City") = "Berlin"
    ctxPanel("header")("Customer.Country") = "DE"
    
    Dim colItems As Collection: Set colItems = New Collection
    Dim i As Long
    For i = 1 To 10
        Dim it As Object: Set it = CreateObject("Scripting.Dictionary")
        it("Items[i].Date") = "2025-09-" & Format(i, "00")
        it("Items[i].Ref") = "REF-" & Format(i, "000")
        it("Items[i].Amount") = 100 + i
        colItems.Add it
    Next i
    Set ctxPanel("repeaters")("Items") = colItems   ' -> rep_Panel_Items
    
    Set root("Panel") = ctxPanel
    
    ' Block: Invoice2 (klassischer Repeater)
    Dim ctxInv As Object: Set ctxInv = oDict_f_p_BuildEmptyContext()
    ctxInv("header")("Invoice.Number") = "INV-2025-092"
    ctxInv("header")("Invoice.Date") = "2025-09-25"
    
    Dim colInv As Collection: Set colInv = New Collection
    AddItem2 colInv, "Beratung", 2, 1250#
    AddItem2 colInv, "Workshop", 1, 2200#
    AddItem2 colInv, "Dokumentation", 3, 400#
    Set ctxInv("repeaters")("Items") = colInv     ' -> rep_Invoice2_Items
    
    Dim sumTotal As Double: sumTotal = SumItems2(colInv)
    ctxInv("totals")("Totals.Sum") = sumTotal
    
    Set root("Invoice2") = ctxInv
    
    Set BuildDemoDataForRenderer2 = root
End Function


Private Sub AddItem2(ByRef coll As Collection, ByVal name As String, ByVal qty As Double, ByVal price As Double)
    Dim it As Object: Set it = CreateObject("Scripting.Dictionary")
    it("Items[i].Name") = name
    it("Items[i].Qty") = qty
    it("Items[i].Price") = price
    it("Items[i].Total") = qty * price
    coll.Add it
End Sub

Private Function SumItems2(ByVal coll As Collection) As Double
    Dim i As Long, s As Double
    For i = 1 To coll.Count
        s = s + CDbl(coll(i)("Items[i].Total"))
    Next
    SumItems2 = s
End Function



Private Sub AddItem(ByRef coll As Collection, name As String, qty As Double, price As Double)
    Dim it As Object: Set it = CreateObject("Scripting.Dictionary")
    it("Items[i].Name") = name
    it("Items[i].Qty") = qty
    it("Items[i].Price") = price
    it("Items[i].Total") = qty * price
    coll.Add it
End Sub

Private Function SumItems(ByVal coll As Collection) As Double
    Dim i As Long, s As Double
    For i = 1 To coll.Count
        s = s + CDbl(coll(i)("Items[i].Total"))
    Next
    SumItems = s
End Function



