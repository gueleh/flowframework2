Attribute VB_Name = "f_pM_TemplRendererLite"
' CORE, do not change
'============================================================================================
'   NAME:     f_pM_TemplRendererLite
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
'============================================================================================
Option Explicit
Option Private Module

Private Const s_m_COMPONENT_NAME As String = "f_pM_TemplRendererLite"

Public Function u_f_p_ParseTemplate(ws As Worksheet, namedBlock As String, namedRepeater As String) As u_f_TRlite_BlockSpec
    Dim rngBlk As Range, rngRep As Range
    Set rngBlk = ws.Range(namedBlock)
    Set rngRep = ws.Range(namedRepeater) ' muss im Block liegen
    
    Dim bs As u_f_TRlite_BlockSpec
    bs.lTop = rngBlk.Row
    bs.lLeft = rngBlk.Column
    bs.lWidth = rngBlk.Columns.Count
    
    Dim r As Long, rowCount As Long
    rowCount = rngBlk.rows.Count
    ReDim bs.uaRowSpecs(1 To rowCount)
    
    Dim repRow As Long: repRow = rngRep.Row
    
    For r = 1 To rowCount
        Dim absRow As Long: absRow = rngBlk.Row + r - 1
        Dim rs As u_f_TRlite_RowSpec
        rs.bIsRepeater = (absRow = repRow)
        
        Dim c As Long, ccount As Long: ccount = rngBlk.Columns.Count
        ReDim rs.uaCellspecs(1 To ccount)
        
        For c = 1 To ccount
            Dim cell As Range
            Set cell = ws.Cells(absRow, rngBlk.Column + c - 1)
            
            Dim cs As u_f_TRlite_CellSpec
            cs.sTemplateText = CStr(cell.Value)
            cs.sPlaceholderList = ExtractPlaceholdersList(cs.sTemplateText) ' "a|b|c" oder ""
            cs.sStyleToken = ExtractStyleToken(cell)
            cs.lRelColInBlock = c
            rs.uaCellspecs(c) = cs
        Next c
        
        bs.uaRowSpecs(r) = rs
    Next r
    
    u_f_p_ParseTemplate = bs
End Function

Private Function ExtractStyleToken(cell As Range) As String
    On Error Resume Next
    Dim cmt As Comment
    Set cmt = cell.Comment
    On Error GoTo 0
    If cmt Is Nothing Then Exit Function
    Dim s As String: s = Trim$(cmt.Text)
    If LCase$(Left$(s, 6)) = "style:" Then ExtractStyleToken = Trim$(Mid$(s, 7))
End Function

Private Function ExtractPlaceholdersList(ByVal s As String) As String
    ' Sucht alle {{...}} und liefert die INHALTE (ohne geschweifte Klammern) als "|"-getrennte Liste
    Dim out As String
    Dim pos As Long, startPos As Long, endPos As Long
    pos = 1
    Do
        startPos = InStr(pos, s, "{{", vbTextCompare)
        If startPos = 0 Then Exit Do
        endPos = InStr(startPos + 2, s, "}}", vbTextCompare)
        If endPos = 0 Then Exit Do
        Dim inner As String
        inner = Mid$(s, startPos + 2, endPos - (startPos + 2))
        If Len(Trim$(inner)) > 0 Then
            If Len(out) > 0 Then out = out & "|"
            out = out & Trim$(inner)
        End If
        pos = endPos + 2
    Loop
    ExtractPlaceholdersList = out ' kann leer sein
End Function

Public Sub WriteBlock(wsOut As Worksheet, ByRef bs As u_f_TRlite_BlockSpec, _
                       header As Object, items As Collection, totals As Object)
    
    Dim outRow As Long: outRow = 1
    Dim outCol As Long: outCol = 1
    
    Dim r As Long
    For r = LBound(bs.uaRowSpecs) To UBound(bs.uaRowSpecs)
        If bs.uaRowSpecs(r).bIsRepeater Then
            Dim k As Long
            For k = 1 To items.Count
                WriteOneRow wsOut, outRow, outCol, bs.lWidth, bs.uaRowSpecs(r), header, items(k), totals
                outRow = outRow + 1
            Next k
        Else
            WriteOneRow wsOut, outRow, outCol, bs.lWidth, bs.uaRowSpecs(r), header, Nothing, totals
            outRow = outRow + 1
        End If
    Next r
End Sub

Private Sub WriteOneRow(wsOut As Worksheet, outRow As Long, outCol As Long, width As Long, _
                        ByRef rs As u_f_TRlite_RowSpec, header As Object, item As Variant, totals As Object)
    Dim vals() As Variant: ReDim vals(1 To 1, 1 To width)
    Dim tokens() As String: ReDim tokens(1 To width)
    
    Dim c As Long
    For c = 1 To width
        Dim cs As u_f_TRlite_CellSpec: cs = rs.uaCellspecs(c)
        
        Dim rendered As String
        rendered = ReplaceAllPlaceholders(cs.sTemplateText, cs.sPlaceholderList, header, item, totals)
        
        vals(1, c) = rendered
        tokens(c) = cs.sStyleToken
    Next c
    
    ' Batch schreiben
    With wsOut.Range(wsOut.Cells(outRow, outCol), wsOut.Cells(outRow, outCol + width - 1))
        .Value = vals
        ApplyStylesToRange .Cells, tokens
    End With
End Sub

Private Function ReplaceAllPlaceholders(ByVal sTemplateText As String, _
                                        ByVal sPlaceholderList As String, _
                                        header As Object, item As Variant, totals As Object) As String
    Dim resultText As String: resultText = sTemplateText
    If Len(sPlaceholderList) = 0 Then
        ReplaceAllPlaceholders = resultText
        Exit Function
    End If
    
    Dim arr() As String: arr = Split(sPlaceholderList, "|")
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        Dim key As String: key = arr(i)
        Dim val As String: val = ResolveValueForKey(key, header, item, totals)
        resultText = Replace(resultText, "{{" & key & "}}", val)
    Next i
    ReplaceAllPlaceholders = resultText
End Function

Private Function ResolveValueForKey(ByVal key As String, header As Object, item As Variant, totals As Object) As String
    On Error Resume Next
    If InStr(1, key, "Items[i].", vbTextCompare) > 0 Then
        If Not IsEmpty(item) Then
            If item.Exists(key) Then ResolveValueForKey = CStr(item(key)) Else ResolveValueForKey = ""
        Else
            ResolveValueForKey = ""
        End If
    ElseIf Left$(key, 7) = "Totals." Then
        If totals.Exists(key) Then ResolveValueForKey = CStr(totals(key)) Else ResolveValueForKey = ""
    Else
        If header.Exists(key) Then ResolveValueForKey = CStr(header(key)) Else ResolveValueForKey = ""
    End If
    On Error GoTo 0
End Function

Private Sub ApplyStylesToRange(rngRow As Range, ByRef styleTokens() As String)
    Dim c As Long
    For c = 1 To rngRow.Columns.Count
        Dim tok As String: tok = styleTokens(c)
        If Len(tok) > 0 Then
            With rngRow.Cells(1, c)
                On Error Resume Next
                .Style = tok
                On Error GoTo 0
                ' Rahmen nach _meta setzen
                ApplyBordersForToken rngRow.Cells(1, c), tok
            End With
        End If
    Next c
End Sub


