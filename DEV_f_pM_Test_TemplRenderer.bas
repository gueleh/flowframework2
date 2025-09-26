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

Public Sub DEV_f_p_TR_RenderInvoiceExample_Legacy()
    f_p_EnsureStylesFromMeta af_wks_Styles.name
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    On Error GoTo CleanFail
    
    Dim wsTpl As Worksheet, wsOut As Worksheet
    Set wsTpl = DEV_f_wks_Template
    Set wsOut = DEV_f_wks_TestCanvas
    wsOut.Cells.Clear
    
    Dim blocks() As BlockSpec2
    blocks = ParseAllBlocks(wsTpl)
    
    Dim data As Object
    Set data = BuildDemoDataForRenderer2()
    
    RenderBlocks wsOut, blocks, data, 1, 1
    
    wsOut.Columns.AutoFit
    
CleanExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub
CleanFail:
   Stop
   Resume
    MsgBox "Renderer2 failed: " & Err.Description, vbCritical
    Resume CleanExit
End Sub

Private Function BuildDemoDataForRenderer2() As Object
    Dim root As Object: Set root = CreateObject("Scripting.Dictionary")
    
    ' Block: Panel (links fix, rechts rep)
    Dim ctxPanel As Object: Set ctxPanel = BuildEmptyContext()
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
    Dim ctxInv As Object: Set ctxInv = BuildEmptyContext()
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

