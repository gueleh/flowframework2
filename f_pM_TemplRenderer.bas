Attribute VB_Name = "f_pM_TemplRenderer"
' CORE, do not change
'============================================================================================
'   NAME:     f_pM_TemplRenderer
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

Private Const s_m_COMPONENT_NAME As String = "f_pM_TemplRenderer"

' ############################################
' # PARSER: BLÖCKE + LANES                    #
' ############################################
Public Function ParseAllBlocks(wsTpl As Worksheet) As BlockSpec2()
    Dim namedRefs() As NamedRangeRef
    namedRefs = CollectNamedRangesInSheet(wsTpl)
    If Not HasElementsNR(namedRefs) Then Err.Raise 5, , "Keine Named Ranges auf '" & wsTpl.name & "'."
    
    ' --- Block-Indices sammeln ---
    Dim blkIdx() As Long
    Dim i As Long
    If HasElementsNR(namedRefs) Then
        For i = LBound(namedRefs) To UBound(namedRefs)
            If LCase$(Left$(namedRefs(i).name, 4)) = "blk_" Then
                AppendIndex blkIdx, i
            End If
        Next i
    End If
    If CountLong(blkIdx) = 0 Then Err.Raise 5, , "Keine blk_* Bereiche gefunden (Template2)."
    
      ' --- Blöcke anlegen ---
      Dim arr() As BlockSpec2
      ReDim arr(1 To CountLong(blkIdx))
      Dim b As Long, idx As Long: idx = 0
      For b = LBound(blkIdx) To UBound(blkIdx)
          Dim br As NamedRangeRef: br = namedRefs(blkIdx(b))
          idx = idx + 1
          With arr(idx)
              .blockKey = Mid$(br.name, 5)   ' nach "blk_"
              Set .ws = br.ws
              .Top = br.rng.Row
              .Left = br.rng.Column
              .RowsCount = br.rng.rows.Count
              .ColsCount = br.rng.Columns.Count
              ' Wichtig: KEIN ReDim (0 To -1) hier!
              ' Optional explizit "leeren": Erase .Lanes
          End With
      Next b

    
    ' --- Blöcke nach Top sortieren ---
    Dim changed As Boolean
    Do
        changed = False
        For b = LBound(arr) To UBound(arr) - 1
            If arr(b).Top > arr(b + 1).Top Then
                Dim t As BlockSpec2: t = arr(b): arr(b) = arr(b + 1): arr(b + 1) = t
                changed = True
            End If
        Next b
    Loop While changed
    
    ' --- Lane-Indices sammeln ---
    Dim laneIdx() As Long
    If HasElementsNR(namedRefs) Then
        For i = LBound(namedRefs) To UBound(namedRefs)
            Dim nm As String: nm = namedRefs(i).name
            If Left$(LCase$(nm), 4) = "fix_" Or Left$(LCase$(nm), 4) = "rep_" Or Left$(LCase$(nm), 4) = "rel_" Then
                AppendIndex laneIdx, i
            End If
        Next i
    End If
    
    ' --- Lanes je Block zuordnen ---
    If CountLong(laneIdx) > 0 Then
        Dim j As Long
        For j = LBound(laneIdx) To UBound(laneIdx)
            Dim lr As NamedRangeRef: lr = namedRefs(laneIdx(j))
            Dim parts() As String: parts = Split(lr.name, "_")
            If UBound(parts) < 2 Then GoTo NextLane
            
            Dim laneType As String: laneType = UCase$(parts(0))          ' FIX | REP | REL
            Dim blockKey As String: blockKey = parts(1)
            Dim laneKey As String: laneKey = JoinMid(parts, 2, "_")
            
            Dim bi As Long: bi = FindBlockIndex(arr, blockKey)
            If bi = 0 Then GoTo NextLane
            
            If Not RangeWithin(arr(bi).ws, lr.rng, arr(bi).Top, arr(bi).Left, arr(bi).RowsCount, arr(bi).ColsCount) Then
                Err.Raise 5, , lr.name & " liegt nicht innerhalb von blk_" & blockKey
            End If
            
            Dim ls As LaneSpec
            ls.laneType = laneType
            ls.key = laneKey
            ls.TopRel = lr.rng.Row - arr(bi).Top + 1
            ls.LeftRel = lr.rng.Column - arr(bi).Left + 1
            ls.RowsCount = lr.rng.rows.Count
            ls.ColsCount = lr.rng.Columns.Count
            ls.Cells = ReadLaneCells(lr.rng)
            ls.PadAfterRows = ExtractPadAfterFromComment(lr.rng.Cells(1, 1)) ' optional
            
            AppendLane arr(bi), ls
NextLane:
        Next j
    End If
    
    ParseAllBlocks = arr
End Function

' ---- Named Ranges im Blatt sammeln (Array, kein Collection) ----
Private Function CollectNamedRangesInSheet(ws As Worksheet) As NamedRangeRef()
    Dim tmp() As NamedRangeRef          ' <-- uninitialisiert lassen!
    Dim n As name
    For Each n In ThisWorkbook.Names
        On Error Resume Next
        Dim r As Range: Set r = n.RefersToRange
        On Error GoTo 0
        If Not r Is Nothing Then
            If r.Worksheet Is ws Then
                Dim nr As NamedRangeRef
                nr.name = n.name
                Set nr.ws = ws
                Set nr.rng = r
                AppendNamedRef tmp, nr      ' AppendNamedRef dimensioniert bei Bedarf
            End If
        End If
    Next n
    CollectNamedRangesInSheet = tmp
End Function


' ---- Array-Append-Helfer ----
Private Sub AppendNamedRef(ByRef arr() As NamedRangeRef, ByRef x As NamedRangeRef)
    Dim L As Long
    On Error Resume Next: L = UBound(arr) + 1: On Error GoTo 0
    If L < 0 Then L = 0
    ReDim Preserve arr(0 To L)
    arr(L) = x
End Sub

Private Sub AppendLane(ByRef blk As BlockSpec2, ByRef ls As LaneSpec)
    Dim L As Long
    On Error Resume Next: L = UBound(blk.lanes) + 1: On Error GoTo 0
    If L < 0 Then L = 0
    ReDim Preserve blk.lanes(0 To L)
    blk.lanes(L) = ls
End Sub

Private Sub AppendIndex(ByRef arr() As Long, ByVal idx As Long)
    Dim L As Long
    On Error Resume Next: L = UBound(arr) + 1: On Error GoTo 0
    If L < 0 Then L = 0
    ReDim Preserve arr(0 To L)
    arr(L) = idx
End Sub

Private Function CountLong(ByRef arr() As Long) As Long
    On Error Resume Next
    CountLong = UBound(arr) - LBound(arr) + 1
    If Err.Number <> 0 Then CountLong = 0: Err.Clear
End Function

Private Function HasElementsNR(ByRef arr() As NamedRangeRef) As Boolean
    On Error Resume Next
    HasElementsNR = (UBound(arr) >= LBound(arr))
    If Err.Number <> 0 Then HasElementsNR = False: Err.Clear
End Function

' ---- Utilities ----
Private Function FindBlockIndex(ByRef blocks() As BlockSpec2, ByVal key As String) As Long
    Dim i As Long
    For i = LBound(blocks) To UBound(blocks)
        If StrComp(blocks(i).blockKey, key, vbTextCompare) = 0 Then
            FindBlockIndex = i
            Exit Function
        End If
    Next i
End Function

Private Function RangeWithin(ws As Worksheet, inner As Range, topAbs As Long, leftAbs As Long, rows As Long, cols As Long) As Boolean
    RangeWithin = Not (inner.Row < topAbs _
                    Or inner.Column < leftAbs _
                    Or inner.Row + inner.rows.Count - 1 > topAbs + rows - 1 _
                    Or inner.Column + inner.Columns.Count - 1 > topAbs + cols - 1)
End Function

Private Function ReadLaneCells(rng As Range) As CellSpec2()
    Dim arr() As CellSpec2
    ReDim arr(1 To rng.rows.Count * rng.Columns.Count)
    Dim idx As Long: idx = 0
    Dim r As Long, c As Long
    For r = 1 To rng.rows.Count
        For c = 1 To rng.Columns.Count
            idx = idx + 1
            Dim cs As CellSpec2
            Dim cell As Range: Set cell = rng.Cells(r, c)
            cs.templateText = CStr(cell.Value)
            cs.PlaceholderList = ExtractPlaceholdersList2(cs.templateText)
            cs.StyleToken = ExtractStyleToken2(cell)
            cs.relRow = r
            cs.relCol = c
            arr(idx) = cs
        Next c
    Next r
    ReadLaneCells = arr
End Function

Private Function ExtractStyleToken2(cell As Range) As String
    On Error Resume Next
    Dim cmt As Comment: Set cmt = cell.Comment
    On Error GoTo 0
    If Not cmt Is Nothing Then
        Dim s As String: s = Trim$(cmt.Text)
        If LCase$(Left$(s, 6)) = "style:" Then
            ExtractStyleToken2 = Trim$(Mid$(s, 7))
        Else
            Dim p As Long: p = InStr(1, LCase$(s), "style:", vbTextCompare)
            If p > 0 Then
                Dim tail As String: tail = Mid$(s, p + 6)
                tail = Split(Split(tail, ";")(0), vbCr)(0)
                ExtractStyleToken2 = Trim$(tail)
            End If
        End If
    End If
End Function

Private Function ExtractPlaceholdersList2(ByVal s As String) As String
    Dim out As String, pos As Long, p1 As Long, p2 As Long
    pos = 1
    Do
        p1 = InStr(pos, s, "{{", vbTextCompare)
        If p1 = 0 Then Exit Do
        p2 = InStr(p1 + 2, s, "}}", vbTextCompare)
        If p2 = 0 Then Exit Do
        Dim inner As String: inner = Trim$(Mid$(s, p1 + 2, p2 - (p1 + 2)))
        If Len(inner) > 0 Then
            If Len(out) > 0 Then out = out & "|"
            out = out & inner
        End If
        pos = p2 + 2
    Loop
    ExtractPlaceholdersList2 = out
End Function

Private Function ExtractPadAfterFromComment(cell As Range) As Long
    On Error Resume Next
    Dim cmt As Comment: Set cmt = cell.Comment
    If cmt Is Nothing Then Exit Function
    Dim s As String: s = cmt.Text
    Dim p As Long: p = InStr(1, LCase$(s), "padafter:", vbTextCompare)
    If p > 0 Then
        Dim tail As String: tail = Mid$(s, p + Len("padafter:"))
        tail = Replace(Replace(tail, vbCr, " "), vbLf, " ")
        tail = Trim$(tail)
        Dim num As String, i As Long
        For i = 1 To Len(tail)
            Dim ch As String: ch = Mid$(tail, i, 1)
            If ch Like "[0-9]" Then
                num = num & ch
            ElseIf Len(num) > 0 Then
                Exit For
            End If
        Next i
        If Len(num) > 0 Then ExtractPadAfterFromComment = CLng(num)
    End If
End Function

Private Function JoinMid(parts() As String, startIdx As Long, sep As String) As String
    Dim i As Long, s As String
    For i = startIdx To UBound(parts)
        If i > startIdx Then s = s & sep
        s = s & parts(i)
    Next i
    JoinMid = s
End Function

' ############################################
' # RENDERING                                 #
' ############################################
Public Sub RenderBlocks(wsOut As Worksheet, ByRef blocks() As BlockSpec2, ByVal data As Object, _
                         ByVal startRowOut As Long, ByVal startColOut As Long)
    Dim outRow As Long: outRow = startRowOut
    Dim outCol As Long: outCol = startColOut
    
    Dim b As Long
    For b = LBound(blocks) To UBound(blocks)
        Dim blk As BlockSpec2: blk = blocks(b)
        Dim ctx As Object
        If data.Exists(blk.blockKey) Then
            Set ctx = data(blk.blockKey)
        Else
            Set ctx = BuildEmptyContext()
        End If
        Dim blockHeight As Long
        blockHeight = RenderOneBlock(wsOut, blk, ctx, outRow, outCol)
        outRow = outRow + blockHeight
    Next b
End Sub

Private Function RenderOneBlock(wsOut As Worksheet, ByRef blk As BlockSpec2, ByVal ctx As Object, _
                                ByVal outTop As Long, ByVal outLeft As Long) As Long
    
      ValidateRepeaters blk, ctx
    
    
    Dim maxBottom As Long: maxBottom = 0
    Dim i As Long
    
    ' 1) FIX zuerst (Template-Position)
    If HasElementsLanes(blk.lanes) Then
        For i = LBound(blk.lanes) To UBound(blk.lanes)
            If UCase$(blk.lanes(i).laneType) = "FIX" Then
                Dim botF As Long
                botF = WriteFixLane(wsOut, blk, blk.lanes(i), ctx, outTop, outLeft)
                If botF > maxBottom Then maxBottom = botF
            End If
        Next i
    End If
    
    ' 2) REP danach (expandiert)
    If HasElementsLanes(blk.lanes) Then
        For i = LBound(blk.lanes) To UBound(blk.lanes)
            If UCase$(blk.lanes(i).laneType) = "REP" Then
                Dim botR As Long
                botR = WriteRepLane(wsOut, blk, blk.lanes(i), ctx, outTop, outLeft)
                If botR > maxBottom Then maxBottom = botR
            End If
        Next i
    End If
    
    ' 3) REL zum Schluss (verschieben, wenn nötig)
    Dim rels() As LaneSpec
    rels = SortRelLanesByTopRel(blk.lanes)
    If HasElementsLanes(rels) Then
        Dim k As Long
        For k = LBound(rels) To UBound(rels)
            Dim botRel As Long
            botRel = WriteRelLane(wsOut, blk, rels(k), ctx, outTop, outLeft, maxBottom)
            If botRel > maxBottom Then maxBottom = botRel
        Next k
    End If
    
    If maxBottom = 0 Then maxBottom = blk.RowsCount
    RenderOneBlock = maxBottom
End Function

Private Function SortRelLanesByTopRel(ByRef lanes() As LaneSpec) As LaneSpec()
    Dim buf() As LaneSpec
    Dim i As Long, n As Long: n = -1
    If HasElementsLanes(lanes) Then
        For i = LBound(lanes) To UBound(lanes)
            If UCase$(lanes(i).laneType) = "REL" Then
                n = n + 1
                ReDim Preserve buf(0 To n)
                buf(n) = lanes(i)
            End If
        Next i
    End If
      If n < 0 Then
          SortRelLanesByTopRel = buf  ' buf ist uninitialisiert ? ok
          Exit Function
      End If

    
    ' sortieren nach TopRel
    If HasElementsLanes(buf) Then
        Dim swapped As Boolean
        Do
            swapped = False
            For i = LBound(buf) To UBound(buf) - 1
                If buf(i).TopRel > buf(i + 1).TopRel Then
                    Dim t As LaneSpec: t = buf(i): buf(i) = buf(i + 1): buf(i + 1) = t
                    swapped = True
                End If
            Next i
        Loop While swapped
    End If
    SortRelLanesByTopRel = buf
End Function

Private Function HasElementsLanes(ByRef lanes() As LaneSpec) As Boolean
    On Error Resume Next
    HasElementsLanes = (UBound(lanes) >= LBound(lanes))
    If Err.Number <> 0 Then HasElementsLanes = False: Err.Clear
End Function

Private Function WriteFixLane(wsOut As Worksheet, ByRef blk As BlockSpec2, ByRef ln As LaneSpec, _
                              ByVal ctx As Object, ByVal outTop As Long, ByVal outLeft As Long) As Long
    Dim r As Long, c As Long
    For r = 1 To ln.RowsCount
        Dim vals() As Variant: ReDim vals(1 To 1, 1 To ln.ColsCount)
        Dim toks() As String: ReDim toks(1 To ln.ColsCount)
        For c = 1 To ln.ColsCount
            Dim cs As CellSpec2: cs = FindCellInLane(ln, r, c)
            vals(1, c) = ReplaceAll2(cs.templateText, cs.PlaceholderList, ctx, Nothing)
            toks(c) = cs.StyleToken
        Next c
        With wsOut.Range( _
            wsOut.Cells(outTop + ln.TopRel + r - 1, outLeft + ln.LeftRel), _
            wsOut.Cells(outTop + ln.TopRel + r - 1, outLeft + ln.LeftRel + ln.ColsCount - 1))
            .Value = vals
            ApplyStylesRow2 .Cells, toks
        End With
    Next r
    Dim bottomRel As Long: bottomRel = ln.TopRel + ln.RowsCount - 1
    If ln.PadAfterRows > 0 Then
        bottomRel = WritePadRows(wsOut, outTop, outLeft, ln, bottomRel)
    End If
    WriteFixLane = bottomRel
End Function

Private Function WriteRepLane(wsOut As Worksheet, ByRef blk As BlockSpec2, ByRef ln As LaneSpec, _
                              ByVal ctx As Object, ByVal outTop As Long, ByVal outLeft As Long) As Long
    Dim items As Collection: Set items = ResolveRepCollection(ctx, ln.key)
    Dim n As Long
    If items Is Nothing Then
      n = 0
   Else
      n = items.Count
   End If
    'n = IIf(items Is Nothing, 0, items.Count)
    
    Dim k As Long, r As Long, c As Long
    For k = 1 To n
        For r = 1 To ln.RowsCount
            Dim vals() As Variant: ReDim vals(1 To 1, 1 To ln.ColsCount)
            Dim toks() As String: ReDim toks(1 To ln.ColsCount)
            For c = 1 To ln.ColsCount
                Dim cs As CellSpec2: cs = FindCellInLane(ln, r, c)
                vals(1, c) = ReplaceAll2(cs.templateText, cs.PlaceholderList, ctx, items(k))
                toks(c) = cs.StyleToken
            Next c
            With wsOut.Range( _
                wsOut.Cells(outTop + ln.TopRel + (k - 1) * ln.RowsCount + r - 1, outLeft + ln.LeftRel), _
                wsOut.Cells(outTop + ln.TopRel + (k - 1) * ln.RowsCount + r - 1, outLeft + ln.LeftRel + ln.ColsCount - 1))
                .Value = vals
                ApplyStylesRow2 .Cells, toks
            End With
        Next r
    Next k
    
    Dim bottomRel As Long
    bottomRel = ln.TopRel + Application.Max(n * ln.RowsCount, ln.RowsCount) - 1
    If ln.PadAfterRows > 0 Then
        bottomRel = WritePadRows(wsOut, outTop, outLeft, ln, bottomRel)
    End If
    WriteRepLane = bottomRel
End Function

Private Function WriteRelLane(wsOut As Worksheet, ByRef blk As BlockSpec2, ByRef ln As LaneSpec, _
                              ByVal ctx As Object, ByVal outTop As Long, ByVal outLeft As Long, _
                              ByVal currentMaxBottom As Long) As Long
    Dim effectiveTopRel As Long
    effectiveTopRel = ln.TopRel
    If currentMaxBottom >= effectiveTopRel Then
        effectiveTopRel = currentMaxBottom + 1
    End If
    
    Dim r As Long, c As Long
    For r = 1 To ln.RowsCount
        Dim vals() As Variant: ReDim vals(1 To 1, 1 To ln.ColsCount)
        Dim toks() As String: ReDim toks(1 To ln.ColsCount)
        For c = 1 To ln.ColsCount
            Dim cs As CellSpec2: cs = FindCellInLane(ln, r, c)
            vals(1, c) = ReplaceAll2(cs.templateText, cs.PlaceholderList, ctx, Nothing)
            toks(c) = cs.StyleToken
        Next c
        With wsOut.Range( _
            wsOut.Cells(outTop + effectiveTopRel + r - 1, outLeft + ln.LeftRel), _
            wsOut.Cells(outTop + effectiveTopRel + r - 1, outLeft + ln.LeftRel + ln.ColsCount - 1))
            .Value = vals
            ApplyStylesRow2 .Cells, toks
        End With
    Next r
    
    Dim bottomRel As Long: bottomRel = effectiveTopRel + ln.RowsCount - 1
    If ln.PadAfterRows > 0 Then
        bottomRel = WritePadRowsAt(wsOut, outTop, outLeft, ln, bottomRel + 1, ln.LeftRel)
    End If
    WriteRelLane = bottomRel
End Function

Private Function WritePadRows(wsOut As Worksheet, ByVal outTop As Long, ByVal outLeft As Long, ByRef ln As LaneSpec, _
                              ByVal currentBottomRel As Long) As Long
    Dim nextRowRel As Long: nextRowRel = currentBottomRel + 1
    WritePadRows = WritePadRowsAt(wsOut, outTop, outLeft, ln, nextRowRel, ln.LeftRel)
End Function

Private Function WritePadRowsAt(wsOut As Worksheet, ByVal outTop As Long, ByVal outLeft As Long, ByRef ln As LaneSpec, _
                                ByVal startRowRel As Long, ByVal startColRel As Long) As Long
    Dim s As Long
    For s = 1 To ln.PadAfterRows
        Dim vals() As Variant: ReDim vals(1 To 1, 1 To ln.ColsCount)
        Dim toks() As String: ReDim toks(1 To ln.ColsCount)
        With wsOut.Range( _
            wsOut.Cells(outTop + startRowRel + s - 1, outLeft + startColRel), _
            wsOut.Cells(outTop + startRowRel + s - 1, outLeft + startColRel + ln.ColsCount - 1))
            .Value = vals
            ApplyStylesRow2 .Cells, toks
        End With
    Next s
    WritePadRowsAt = startRowRel + ln.PadAfterRows - 1
End Function

Private Function FindCellInLane(ByRef ln As LaneSpec, ByVal relRow As Long, ByVal relCol As Long) As CellSpec2
    Dim idx As Long: idx = (relRow - 1) * ln.ColsCount + relCol
    FindCellInLane = ln.Cells(idx)
End Function

Private Sub ApplyStylesRow2(rngRow As Range, ByRef styleTokens() As String)
    Dim c As Long
    For c = 1 To rngRow.Columns.Count
        Dim tok As String: tok = styleTokens(c)
        If Len(tok) > 0 Then
            On Error Resume Next
            rngRow.Cells(1, c).Style = tok
            On Error GoTo 0
            ' Rahmen nach Token aus _meta
            ApplyBordersForToken rngRow.Cells(1, c), tok
        End If
    Next c
End Sub


' ############################################
' # PLATZHALTER / DATEN-KONTEXT               #
' ############################################
Private Function ReplaceAll2(ByVal templateText As String, ByVal listKeys As String, _
                             ByVal ctx As Object, ByVal item As Variant) As String
    Dim out As String: out = templateText
    If Len(listKeys) = 0 Then ReplaceAll2 = out: Exit Function
    Dim arr() As String: arr = Split(listKeys, "|")
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        Dim k As String: k = arr(i)
        Dim v As String: v = ResolveValue2(k, ctx, item)
        out = Replace(out, "{{" & k & "}}", v)
    Next i
    ReplaceAll2 = out
End Function

Private Function ResolveValue2(ByVal key As String, ByVal ctx As Object, ByVal item As Variant) As String
    On Error Resume Next
    If InStr(1, key, "[i].", vbTextCompare) > 0 Then
        If Not IsEmpty(item) Then
            If item.Exists(key) Then ResolveValue2 = CStr(item(key)) Else ResolveValue2 = ""
        Else
            ResolveValue2 = ""
        End If
    ElseIf Left$(key, 7) = "Totals." Then
        If ctx("totals").Exists(key) Then ResolveValue2 = CStr(ctx("totals")(key)) Else ResolveValue2 = ""
    Else
        If ctx("header").Exists(key) Then ResolveValue2 = CStr(ctx("header")(key)) Else ResolveValue2 = ""
    End If
    On Error GoTo 0
End Function

Private Function ResolveRepCollection(ByVal ctx As Object, ByVal laneKey As String) As Collection
    On Error Resume Next
    Dim reps As Object: Set reps = ctx("repeaters")
    If Not reps Is Nothing Then
        If reps.Exists(laneKey) Then Set ResolveRepCollection = reps(laneKey)
    End If
    On Error GoTo 0
End Function

Public Function BuildEmptyContext() As Object
    Dim ctx As Object: Set ctx = CreateObject("Scripting.Dictionary")
    Set ctx("header") = CreateObject("Scripting.Dictionary")
    Set ctx("totals") = CreateObject("Scripting.Dictionary")
    Set ctx("repeaters") = CreateObject("Scripting.Dictionary")
    Set BuildEmptyContext = ctx
End Function

Private Sub ValidateRepeaters(ByRef blk As BlockSpec2, ByVal ctx As Object)
    Dim i As Long
    For i = LBound(blk.lanes) To UBound(blk.lanes)
        If UCase$(blk.lanes(i).laneType) = "REP" Then
            If Not ctx("repeaters").Exists(blk.lanes(i).key) Then
                Err.Raise 5, , "Fehlende Repeater-Daten: blk_" & blk.blockKey & _
                               " ? rep key '" & blk.lanes(i).key & "'"
            End If
        End If
    Next i
End Sub


