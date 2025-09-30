Attribute VB_Name = "f_pM_TemplRenderer_Styles"
' CORE, do not change
'============================================================================================
'   NAME:     f_pM_TemplRenderer_Styles
'============================================================================================
'   Purpose:  processing the styles for Template Renderer and Template Renderer Lite
'   Access:   Private
'   Type:     Module
'   Author:   Günther Lehner
'   Contact:  guenther.lehner@protonmail.com
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
' TODO: [+] Integrate into framework
' TODO: [+] Dictionaries auf Nothing setzen, sobald sie nicht mehr gebraucht werden
'============================================================================================
Option Explicit
Option Private Module

Private Const s_m_COMPONENT_NAME As String = "f_pM_TemplRenderer_Styles"

Public gStyleBorderSpec As Object        ' token -> Dict of edges ("OUTLINE","TOP",...)
Public gStyleBorderWeight As Object      ' token -> XlBorderWeight

' ------------------------------------------------------------------
'  EnsureStylesFromMeta "_meta"
'  Liest _meta!Styles (benannter Bereich) und baut/aktualisiert Styles.
'  Benötigte Spaltennamen (Groß/Klein egal):
'  Token | NumberFormat | HAlign | VAlign | Wrap | Indent | FontName
'  FontSize | Bold | Italic | FontColor | FillColor | BorderSpec | BorderWeight
' ------------------------------------------------------------------
Public Sub f_p_EnsureStylesFromMeta(metaSheetName As String, Optional ByRef oWkb_arg_Destination As Workbook)
' TODO: Refactor for framework compliance
    Dim ws As Worksheet, rng As Range
    Dim oWkbTarget As Workbook
    
    Set oWkbTarget = oWkb_f_p_DefaultToThisWorkbook(oWkb_arg_Destination)
    
    Set ws = ThisWorkbook.Worksheets(metaSheetName)
    
        Dim token As String
        Dim st As Style
    
    On Error Resume Next
    Set rng = ws.Range("af_rng_Styles")       ' benannter Bereich "Styles"
    On Error GoTo 0
    If rng Is Nothing Then Err.Raise 5, , metaSheetName & ": benannter Bereich 'af_rng_Styles' fehlt."

    ' Header-Map (Name -> Spaltenindex im rng)
    Dim H As Object: Set H = CreateObject("Scripting.Dictionary")
    Dim j As Long
    For j = 1 To rng.Columns.Count
        Dim key As String
        key = LCase$(Trim$(CStr(rng.Cells(1, j).Value)))
        If Len(key) > 0 Then H(key) = j
    Next j
    
    If gStyleBorderSpec Is Nothing Then Set gStyleBorderSpec = CreateObject("Scripting.Dictionary")
    If gStyleBorderWeight Is Nothing Then Set gStyleBorderWeight = CreateObject("Scripting.Dictionary")
    gStyleBorderSpec.RemoveAll
    gStyleBorderWeight.RemoveAll
    
    Dim r As Long
    For r = 2 To rng.rows.Count
        token = GetText(rng, H, r, "token")
        If Len(token) = 0 Then GoTo NextRow
        
         Set st = Nothing
        On Error Resume Next
        Set st = oWkbTarget.Styles(token)
        On Error GoTo 0
        If st Is Nothing Then Set st = oWkbTarget.Styles.Add(token)
        
        ' Number format
        Dim nf As String: nf = GetText(rng, H, r, "numberformat")
        If Len(nf) > 0 Then st.NumberFormat = nf
        
        ' Alignment
        Dim ha As Variant: ha = MapHAlign(GetText(rng, H, r, "halign"))
        If Not IsEmpty(ha) Then st.HorizontalAlignment = ha
        Dim va As Variant: va = MapVAlign(GetText(rng, H, r, "valign"))
        If Not IsEmpty(va) Then st.VerticalAlignment = va
        
        Dim wrap As Variant: wrap = GetBool(rng, H, r, "wrap")
        If Not IsEmpty(wrap) Then st.WrapText = wrap
        
        ' Indent (Style unterstützt es – dennoch mit Fehlerfang)
        Dim ind As Variant: ind = GetLong(rng, H, r, "indent")
        If Not IsEmpty(ind) Then On Error Resume Next: st.IndentLevel = ind: On Error GoTo 0
        
        ' Font
        Dim fn As String: fn = GetText(rng, H, r, "fontname"): If Len(fn) > 0 Then st.Font.name = fn
        Dim fs As Variant: fs = GetDouble(rng, H, r, "fontsize"): If Not IsEmpty(fs) Then st.Font.Size = fs
        Dim fb As Variant: fb = GetBool(rng, H, r, "bold"): If Not IsEmpty(fb) Then st.Font.Bold = fb
        Dim fi As Variant: fi = GetBool(rng, H, r, "italic"): If Not IsEmpty(fi) Then st.Font.Italic = fi
        
        Dim fcol As Variant: fcol = ParseColor(GetText(rng, H, r, "fontcolor"))
        If Not IsEmpty(fcol) Then st.Font.Color = fcol
        Dim bcol As Variant: bcol = ParseColor(GetText(rng, H, r, "fillcolor"))
        If Not IsEmpty(bcol) Then st.Interior.Color = bcol
        
        ' Style-Teilbereiche explizit aktivieren
        st.IncludeNumber = True
        st.IncludeFont = True
        st.IncludeAlignment = True
        st.IncludePatterns = True
        
        ' Borders: nur merken – angewendet wird später je Zelle
        Dim bspec As String: bspec = UCase$(GetText(rng, H, r, "borderspec"))
        If Len(bspec) > 0 Then
            Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
            Dim parts() As String
            parts = Split(Replace(Replace(bspec, ";", ","), "|", ","), ",")
            For j = LBound(parts) To UBound(parts)
                Dim t As String: t = Trim$(parts(j))
                If Len(t) > 0 Then d(t) = True
            Next j
            Set gStyleBorderSpec(token) = d
        End If
        
        Dim w As Long: w = MapBorderWeight(GetText(rng, H, r, "borderweight"))
        gStyleBorderWeight(token) = w
        
NextRow:
    Next r
End Sub

' ---------- Öffentliche Border-Anwendung ----------
Public Sub ApplyBordersForToken(ByVal target As Range, ByVal token As String)
    If gStyleBorderSpec Is Nothing Then Exit Sub
    If Not gStyleBorderSpec.Exists(token) Then Exit Sub
' TODO: Refactor for framework compliance
    Dim w As Long: w = xlThin
    If Not gStyleBorderWeight Is Nothing Then
        If gStyleBorderWeight.Exists(token) Then w = gStyleBorderWeight(token)
    End If
    
    Dim spec As Object: Set spec = gStyleBorderSpec(token)
    
    ' 1) Kanten, die wir kontrollieren, erst mal löschen
    ClearManagedBorders target
    
    ' 2) Spezifizierte Kanten setzen
    If spec.Exists("OUTLINE") Then
        ApplyEdge target, xlEdgeTop, w
        ApplyEdge target, xlEdgeBottom, w
        ApplyEdge target, xlEdgeLeft, w
        ApplyEdge target, xlEdgeRight, w
    End If
    If spec.Exists("TOP") Then ApplyEdge target, xlEdgeTop, w
    If spec.Exists("BOTTOM") Then ApplyEdge target, xlEdgeBottom, w
    If spec.Exists("LEFT") Then ApplyEdge target, xlEdgeLeft, w
    If spec.Exists("RIGHT") Then ApplyEdge target, xlEdgeRight, w
    ' INSIDE funktioniert sinnvoll nur auf Mehrzellen-Ranges
    If target.Cells.CountLarge > 1 Then
        If spec.Exists("INSIDEH") Then ApplyInside target, True, w
        If spec.Exists("INSIDEV") Then ApplyInside target, False, w
    End If
End Sub

Private Sub ClearManagedBorders(ByVal rng As Range)
' TODO: Refactor for framework compliance
    Dim idx As Variant
    For Each idx In Array(xlEdgeTop, xlEdgeBottom, xlEdgeLeft, xlEdgeRight, xlInsideHorizontal, xlInsideVertical)
        With rng.Borders(idx)
            .LineStyle = xlNone
        End With
    Next idx
End Sub

Private Sub ApplyEdge(ByVal rng As Range, ByVal edge As XlBordersIndex, ByVal weight As XlBorderWeight)
' TODO: Refactor for framework compliance
    With rng.Borders(edge)
        .LineStyle = xlContinuous
        .weight = weight
        ' Farbe bleibt automatisch (Automatisch); bei Bedarf:
        ' .ColorIndex = xlColorIndexAutomatic
    End With
End Sub

Private Sub ApplyInside(ByVal rng As Range, ByVal horizontal As Boolean, ByVal weight As XlBorderWeight)
' TODO: Refactor for framework compliance
    Dim idx As XlBordersIndex
    If horizontal Then idx = xlInsideHorizontal Else idx = xlInsideVertical
    With rng.Borders(idx)
        .LineStyle = xlContinuous
        .weight = weight
    End With
End Sub


' Kompatibilität – falls Altcode sie noch aufruft
Public Sub ApplyOutlineBorderToCell(ByVal target As Range)
    ApplyEdge target, xlEdgeTop, xlThin
    ApplyEdge target, xlEdgeBottom, xlThin
    ApplyEdge target, xlEdgeLeft, xlThin
    ApplyEdge target, xlEdgeRight, xlThin
End Sub

Private Function MapBorderWeight(ByVal s As String) As Long
' TODO: Refactor for framework compliance
    s = UCase$(Trim$(s))
    Select Case s
        Case "", "THIN", "XLTHIN": MapBorderWeight = xlThin
        Case "MEDIUM", "XLMEDIUM": MapBorderWeight = xlMedium
        Case "THICK", "XLTHICK": MapBorderWeight = xlThick
        Case Else
            If IsNumeric(s) Then MapBorderWeight = CLng(s) Else MapBorderWeight = xlThin
    End Select
End Function

Private Function MapHAlign(ByVal s As String) As Variant
' TODO: Refactor for framework compliance
    s = UCase$(Trim$(s))
    Select Case s
        Case "LEFT": MapHAlign = xlHAlignLeft
        Case "CENTER", "CENTRE", "MIDDLE": MapHAlign = xlHAlignCenter
        Case "RIGHT": MapHAlign = xlHAlignRight
        Case Else: MapHAlign = Empty
    End Select
End Function

Private Function MapVAlign(ByVal s As String) As Variant
' TODO: Refactor for framework compliance
    s = UCase$(Trim$(s))
    Select Case s
        Case "TOP": MapVAlign = xlVAlignTop
        Case "CENTER", "CENTRE", "MIDDLE": MapVAlign = xlVAlignCenter
        Case "BOTTOM": MapVAlign = xlVAlignBottom
        Case Else: MapVAlign = Empty
    End Select
End Function

Private Function GetText(rng As Range, H As Object, rowIdx As Long, key As String) As String
' TODO: Refactor for framework compliance
    key = LCase$(key)
    If Not H.Exists(key) Then Exit Function
    GetText = Trim$(CStr(rng.Cells(rowIdx, H(key)).Value))
End Function

Private Function GetBool(rng As Range, H As Object, rowIdx As Long, key As String) As Variant
' TODO: Refactor for framework compliance
    key = LCase$(key)
    If Not H.Exists(key) Then Exit Function
    Dim v As Variant: v = rng.Cells(rowIdx, H(key)).Value
    If VarType(v) = vbBoolean Then GetBool = CBool(v): Exit Function
    Dim s As String: s = UCase$(Trim$(CStr(v)))
    If s = "TRUE" Or s = "JA" Or s = "1" Then GetBool = True: Exit Function
    If s = "FALSE" Or s = "NEIN" Or s = "0" Then GetBool = False: Exit Function
    GetBool = Empty
End Function

Private Function GetLong(rng As Range, H As Object, rowIdx As Long, key As String) As Variant
' TODO: Refactor for framework compliance
    key = LCase$(key)
    If Not H.Exists(key) Then Exit Function
    Dim v As Variant: v = rng.Cells(rowIdx, H(key)).Value
    If IsNumeric(v) Then GetLong = CLng(v) Else GetLong = Empty
End Function

Private Function GetDouble(rng As Range, H As Object, rowIdx As Long, key As String) As Variant
' TODO: Refactor for framework compliance
    key = LCase$(key)
    If Not H.Exists(key) Then Exit Function
    Dim v As Variant: v = rng.Cells(rowIdx, H(key)).Value
    If IsNumeric(v) Then GetDouble = CDbl(v) Else GetDouble = Empty
End Function

' Akzeptiert: leere Zelle -> Empty,
' Zahl (Long) -> direkt, "#RRGGBB" -> HTML-Hex, "R,G,B" -> CSV
Private Function ParseColor(ByVal s As String) As Variant
    s = Trim$(s)
    If Len(s) = 0 Then Exit Function
    If IsNumeric(s) Then ParseColor = CLng(s): Exit Function
    
    If Left$(s, 1) = "#" And Len(s) = 7 Then
        Dim r As Long, g As Long, b As Long
        r = CLng("&H" & Mid$(s, 2, 2))
        g = CLng("&H" & Mid$(s, 4, 2))
        b = CLng("&H" & Mid$(s, 6, 2))
        ParseColor = RGB(r, g, b)
        Exit Function
    End If
    
    Dim parts() As String
    parts = Split(s, ",")
    If UBound(parts) = 2 Then
        ParseColor = RGB(val(parts(0)), val(parts(1)), val(parts(2)))
    End If
End Function



