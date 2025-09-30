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
'   Usage: please refer to DEV_f_pM_Test_TemplRenderer to see how this is used
'--------------------------------------------------------------------------------------------
'   VERSION HISTORY
'   Version    Date    Developer    Changes
'   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'--------------------------------------------------------------------------------------------
'   BACKLOG
'   ''''''''''''''''''''
' TODO: [+] Integrate into framework
' TODO: PadAfter does not work for rel_ lanes
'============================================================================================
Option Explicit
Option Private Module

Private Const s_m_COMPONENT_NAME As String = "f_pM_TemplRenderer"

Private Const s_m_BLOCK_PREFIX As String = "blk_"
Private Const s_m_FIXED_LANE_PREFIX As String = "fix_"
Private Const s_m_REPEATER_LANE_PREFIX As String = "rep_"
Private Const s_m_RELATIVE_LANE_PREFIX As String = "rel_"
Private Const s_m_LANE_TAG_FIX As String = "FIX"
Private Const s_m_LANE_TAG_REPEATABLE As String = "REP"
Private Const s_m_LANE_TAG_RELATIVE As String = "REL"
Private Const s_m_KEY_HEADER As String = "header"
Private Const s_m_KEY_REPEATERS As String = "repeaters"
Private Const s_m_KEY_TOTALS As String = "totals"
Private Const s_m_SEPARATOR As String = "|"

' Purpose: sets up top level dictionary for template structure
Public Function oDict_f_p_BuildEmptyContext() As Scripting.Dictionary
    Dim oDictContext As Scripting.Dictionary
    Set oDictContext = New Scripting.Dictionary
    Set oDictContext(s_m_KEY_HEADER) = New Scripting.Dictionary
    Set oDictContext(s_m_KEY_TOTALS) = New Scripting.Dictionary
    Set oDictContext(s_m_KEY_REPEATERS) = New Scripting.Dictionary
    Set oDict_f_p_BuildEmptyContext = oDictContext
End Function

' Purpose: parse template from provided sheet and return block specs and in them all other relevant specs for the renderer
Public Function b_f_p_GetAllParsedBlockSpecs( _
   ByRef ua_arg_ParsedBlockSpecs() As u_f_BlockSpecRenderer, _
   ByRef oWksTemplate As Worksheet _
) As Boolean

'Fixed, don't change
   Dim oC_Me As New f_C_CallParams: oC_Me.s_prop_rw_ComponentName = s_m_COMPONENT_NAME: If oC_f_p_FrameworkSettings.b_prop_rw_ThisIsATestRun Then f_p_RegisterUnitTest oC_Me
'>>>>>>> Your custom settings here
   With oC_Me
      .s_prop_rw_ProcedureName = "b_f_p_GetAllParsedBlockSpecs" 'Name of the function
      .b_prop_rw_SilentError = True 'False will display a message box - you should only do this on entry level
      .s_prop_rw_ErrorMessage = "Parsing of BlockSpecs from Template failed." 'A message that properly informs the user and the devs (silent errors will be logged nonetheless)
      .SetCallArgs "No args" 'If the sub takes args put the here like ("sExample:=" & sExample, "lExample:=" & lExample)
   End With
'Fixed, don't change
Try: On Error GoTo Catch


'>>>>>>> Your code here
   Dim uaNamedRefs() As u_f_NamedRangeRefRenderer
   Dim laBlockIndex() As Long
   Dim lIndex As Long
   Dim uaBlockSpecs() As u_f_BlockSpecRenderer
   Dim lBlockIndex As Long
   Dim lIndexBlockSpec As Long
   Dim laLaneIndex() As Long
   Dim bChanged As Boolean
   Dim sRngName As String
   Dim lLaneIndex As Long
   Dim saParts() As String
   Dim sLaneType As String
   Dim sBlockKey As String
   Dim sLaneKey As String
    
   uaNamedRefs = ua_m_CollectNamedRangesSpecsFromSheet(oWksTemplate)
   If Not b_m_HasElementsNR(uaNamedRefs) Then Err.Raise 5, , "Keine Named Ranges auf '" & oWksTemplate.name & "'."
   
   ' --- Block-Indices sammeln ---
   If b_m_HasElementsNR(uaNamedRefs) Then
      For lIndex = LBound(uaNamedRefs) To UBound(uaNamedRefs)
         If LCase$(Left$(uaNamedRefs(lIndex).sName, 4)) = s_m_BLOCK_PREFIX Then
            mAppendIndex laBlockIndex, lIndex
         End If
      Next lIndex
   End If
   If l_m_CountLong(laBlockIndex) = 0 Then Err.Raise 5, , "Keine blk_* Bereiche gefunden (Template2)."
    
   ' --- Blöcke anlegen ---
   ReDim uaBlockSpecs(1 To l_m_CountLong(laBlockIndex))
   lIndexBlockSpec = 0
   For lBlockIndex = LBound(laBlockIndex) To UBound(laBlockIndex)
      Dim uNamedRangeSpec As u_f_NamedRangeRefRenderer
      uNamedRangeSpec = uaNamedRefs(laBlockIndex(lBlockIndex))
      lIndexBlockSpec = lIndexBlockSpec + 1
      With uaBlockSpecs(lIndexBlockSpec)
         .sBlockKey = Mid$(uNamedRangeSpec.sName, 5)   ' nach "blk_"
         Set .oWks = uNamedRangeSpec.oWks
         .lTop = uNamedRangeSpec.oRng.Row
         .lLeft = uNamedRangeSpec.oRng.Column
         .lRowsCount = uNamedRangeSpec.oRng.rows.Count
         .lColsCount = uNamedRangeSpec.oRng.Columns.Count
      End With
   Next lBlockIndex

    
   ' --- Blöcke nach Top sortieren ---
   Do
      bChanged = False
      For lBlockIndex = LBound(uaBlockSpecs) To UBound(uaBlockSpecs) - 1
         If uaBlockSpecs(lBlockIndex).lTop > uaBlockSpecs(lBlockIndex + 1).lTop Then
            Dim uTemp As u_f_BlockSpecRenderer
            uTemp = uaBlockSpecs(lBlockIndex)
            uaBlockSpecs(lBlockIndex) = uaBlockSpecs(lBlockIndex + 1)
            uaBlockSpecs(lBlockIndex + 1) = uTemp
            bChanged = True
         End If
      Next lBlockIndex
   Loop While bChanged
    
   ' --- Lane-Indices sammeln ---
   If b_m_HasElementsNR(uaNamedRefs) Then
      For lIndex = LBound(uaNamedRefs) To UBound(uaNamedRefs)
         sRngName = uaNamedRefs(lIndex).sName
         If Left$(LCase$(sRngName), 4) = s_m_FIXED_LANE_PREFIX _
         Or Left$(LCase$(sRngName), 4) = s_m_REPEATER_LANE_PREFIX _
         Or Left$(LCase$(sRngName), 4) = s_m_RELATIVE_LANE_PREFIX _
         Then
            mAppendIndex laLaneIndex, lIndex
         End If
      Next lIndex
   End If
    
   ' --- Lanes je Block zuordnen ---
   If l_m_CountLong(laLaneIndex) > 0 Then
      For lLaneIndex = LBound(laLaneIndex) To UBound(laLaneIndex)
         Dim uLaneRangeSpec As u_f_NamedRangeRefRenderer
         uLaneRangeSpec = uaNamedRefs(laLaneIndex(lLaneIndex))
         saParts = Split(uLaneRangeSpec.sName, "_")
         If UBound(saParts) < 2 Then GoTo NextLane
         
         sLaneType = UCase$(saParts(0))          ' FIX | REP | REL
         sBlockKey = saParts(1)
         sLaneKey = s_m_JoinMid(saParts, 2, "_")
         
         lBlockIndex = l_m_FindBlockIndex(uaBlockSpecs, sBlockKey)
         If lBlockIndex = 0 Then GoTo NextLane
         
         If Not b_m_RangeWithin( _
               uaBlockSpecs(lBlockIndex).oWks, _
               uLaneRangeSpec.oRng, _
               uaBlockSpecs(lBlockIndex).lTop, _
               uaBlockSpecs(lBlockIndex).lLeft, _
               uaBlockSpecs(lBlockIndex).lRowsCount, _
               uaBlockSpecs(lBlockIndex).lColsCount _
            ) _
         Then
            Err.Raise 5, , uLaneRangeSpec.sName & " liegt nicht innerhalb von blk_" & sBlockKey
         End If
          
         Dim uLaneSpec As u_f_LaneSpecRenderer
         With uLaneSpec
            .sLaneType = sLaneType
            .sKey = sLaneKey
            .lTopRel = uLaneRangeSpec.oRng.Row - uaBlockSpecs(lBlockIndex).lTop
            .lLeftRel = uLaneRangeSpec.oRng.Column - uaBlockSpecs(lBlockIndex).lLeft
            .lRowsCount = uLaneRangeSpec.oRng.rows.Count
            .lColsCount = uLaneRangeSpec.oRng.Columns.Count
            .uaCells = ua_m_ReadLaneCellSpecs(uLaneRangeSpec.oRng)
            .lPadAfterRows = l_m_ExtractPadAfterFromComment(uLaneRangeSpec.oRng.Cells(1, 1)) ' optional
         End With
         mAppendLane uaBlockSpecs(lBlockIndex), uLaneSpec
NextLane:
      Next lLaneIndex
   End If 'If CountLong(laLaneIndex) > 0 Then
    
   ua_arg_ParsedBlockSpecs = uaBlockSpecs



'End of your code <<<<<<<


'Fixed, don't change
Finally: On Error Resume Next


'>>>>>>> Your code here



'End of your code <<<<<<<


'>>>>>>> Your custom settings here
   If oC_Me.oC_prop_r_Error Is Nothing Then b_f_p_GetAllParsedBlockSpecs = True 'reports execution as successful to caller
'Fixed, don't change
   Exit Function
HandleError: af_pM_ErrorHandling.af_p_Hook_ErrorHandling_LowerLevel


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
End Function


' ---- Named Ranges im Blatt sammeln (Array, kein Collection) ----
Private Function ua_m_CollectNamedRangesSpecsFromSheet(ByRef oWks As Worksheet) As u_f_NamedRangeRefRenderer()
   Dim uaTmpNamdRangeSpec() As u_f_NamedRangeRefRenderer          ' <-- uninitialisiert lassen!
   Dim oName As name
   Dim oRng As Range
   
   For Each oName In ThisWorkbook.Names
      On Error Resume Next
      Set oRng = oName.RefersToRange
      On Error GoTo 0
      If Not oRng Is Nothing Then
         If oRng.Worksheet Is oWks Then
            Dim uNamedRangeSpec As u_f_NamedRangeRefRenderer
            uNamedRangeSpec.sName = oName.name
            Set uNamedRangeSpec.oWks = oWks
            Set uNamedRangeSpec.oRng = oRng
            mAppendNamedRef uaTmpNamdRangeSpec, uNamedRangeSpec      ' AppendNamedRef dimensioniert bei Bedarf
         End If
      End If
   Next oName
   ua_m_CollectNamedRangesSpecsFromSheet = uaTmpNamdRangeSpec
End Function


' ---- Array-Append-Helfer ----
Private Sub mAppendNamedRef(ByRef arr() As u_f_NamedRangeRefRenderer, ByRef x As u_f_NamedRangeRefRenderer)
' TODO: refactor for coding style compliance
    Dim L As Long
    On Error Resume Next: L = UBound(arr) + 1: On Error GoTo 0
    If L < 0 Then L = 0
    ReDim Preserve arr(0 To L)
    arr(L) = x
End Sub

Private Sub mAppendLane(ByRef uBlockSpec As u_f_BlockSpecRenderer, ByRef uLaneSpec As u_f_LaneSpecRenderer)
    Dim lIndex As Long
    On Error Resume Next
    lIndex = UBound(uBlockSpec.uaLanes) + 1
    On Error GoTo 0
    If lIndex < 0 Then lIndex = 0
    ReDim Preserve uBlockSpec.uaLanes(0 To lIndex)
    uBlockSpec.uaLanes(lIndex) = uLaneSpec
End Sub

Private Sub mAppendIndex(ByRef arr() As Long, ByVal idx As Long)
' TODO: refactor for coding style compliance
    Dim L As Long
    On Error Resume Next: L = UBound(arr) + 1: On Error GoTo 0
    If L < 0 Then L = 0
    ReDim Preserve arr(0 To L)
    arr(L) = idx
End Sub

Private Function l_m_CountLong(ByRef arr() As Long) As Long
' TODO: refactor for coding style compliance
    On Error Resume Next
    l_m_CountLong = UBound(arr) - LBound(arr) + 1
    If Err.Number <> 0 Then l_m_CountLong = 0: Err.Clear
End Function

Private Function b_m_HasElementsNR(ByRef arr() As u_f_NamedRangeRefRenderer) As Boolean
' TODO: refactor for coding style compliance
    On Error Resume Next
    b_m_HasElementsNR = (UBound(arr) >= LBound(arr))
    If Err.Number <> 0 Then b_m_HasElementsNR = False: Err.Clear
End Function

' ---- Utilities ----
Private Function l_m_FindBlockIndex(ByRef blocks() As u_f_BlockSpecRenderer, ByVal key As String) As Long
' TODO: refactor for coding style compliance
    Dim i As Long
    For i = LBound(blocks) To UBound(blocks)
        If StrComp(blocks(i).sBlockKey, key, vbTextCompare) = 0 Then
            l_m_FindBlockIndex = i
            Exit Function
        End If
    Next i
End Function

Private Function b_m_RangeWithin(ws As Worksheet, inner As Range, topAbs As Long, leftAbs As Long, rows As Long, cols As Long) As Boolean
    b_m_RangeWithin = Not (inner.Row < topAbs _
                    Or inner.Column < leftAbs _
                    Or inner.Row + inner.rows.Count - 1 > topAbs + rows - 1 _
                    Or inner.Column + inner.Columns.Count - 1 > topAbs + cols - 1)
End Function

Private Function ua_m_ReadLaneCellSpecs(rng As Range) As u_f_CellSpecRenderer()
' TODO: refactor for coding style compliance
    Dim arr() As u_f_CellSpecRenderer
    ReDim arr(1 To rng.rows.Count * rng.Columns.Count)
    Dim idx As Long: idx = 0
    Dim r As Long, c As Long
    For r = 1 To rng.rows.Count
        For c = 1 To rng.Columns.Count
            idx = idx + 1
            Dim cs As u_f_CellSpecRenderer
            Dim cell As Range: Set cell = rng.Cells(r, c)
            cs.sTemplateText = CStr(cell.Value)
            cs.sPlaceholderList = s_m_ExtractPlaceholdersList(cs.sTemplateText)
            cs.sStyleToken = s_m_ExtractStyleToken(cell)
            cs.lRelRow = r
            cs.lRelCol = c
            arr(idx) = cs
        Next c
    Next r
    ua_m_ReadLaneCellSpecs = arr
End Function

Private Function s_m_ExtractStyleToken(cell As Range) As String
' TODO: refactor for coding style compliance
    On Error Resume Next
    Dim cmt As Comment: Set cmt = cell.Comment
    On Error GoTo 0
    If Not cmt Is Nothing Then
        Dim s As String: s = Trim$(cmt.Text)
        If LCase$(Left$(s, 6)) = "style:" Then
            s_m_ExtractStyleToken = Trim$(Mid$(s, 7))
        Else
            Dim p As Long: p = InStr(1, LCase$(s), "style:", vbTextCompare)
            If p > 0 Then
                Dim tail As String: tail = Mid$(s, p + 6)
                tail = Split(Split(tail, ";")(0), vbCr)(0)
                s_m_ExtractStyleToken = Trim$(tail)
            End If
        End If
    End If
End Function

Private Function s_m_ExtractPlaceholdersList(ByVal s As String) As String
' TODO: refactor for coding style compliance
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
    s_m_ExtractPlaceholdersList = out
End Function

Private Function l_m_ExtractPadAfterFromComment(cell As Range) As Long
' TODO: refactor for coding style compliance
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
        If Len(num) > 0 Then l_m_ExtractPadAfterFromComment = CLng(num)
    End If
End Function

Private Function s_m_JoinMid(ByRef saParts() As String, ByVal lStartIndex As Long, ByVal sSeparator As String) As String
   Dim lIndex As Long, sTemp As String
   For lIndex = lStartIndex To UBound(saParts)
      If lIndex > lStartIndex Then sTemp = sTemp & sSeparator
      sTemp = sTemp & saParts(lIndex)
   Next lIndex
   s_m_JoinMid = sTemp
End Function

' Purpose: renders the blocks in the provided output sheet from the provided position, based on block specs and data to replace the tags in it
Public Function b_f_p_RenderBlocks( _
   ByRef oWksOut As Worksheet, _
   ByRef uaBlockSpecs() As u_f_BlockSpecRenderer, _
   ByVal oDictBlocksData As Scripting.Dictionary, _
   ByVal lStartRowOut As Long, _
   ByVal lStartColOut As Long, _
   Optional ByRef lNextStartRow As Long _
) As Boolean

'Fixed, don't change
   Dim oC_Me As New f_C_CallParams: oC_Me.s_prop_rw_ComponentName = s_m_COMPONENT_NAME: If oC_f_p_FrameworkSettings.b_prop_rw_ThisIsATestRun Then f_p_RegisterUnitTest oC_Me
'>>>>>>> Your custom settings here
   With oC_Me
      .s_prop_rw_ProcedureName = "b_f_p_RenderBlocks" 'Name of the function
      .b_prop_rw_SilentError = True 'False will display a message box - you should only do this on entry level
      .s_prop_rw_ErrorMessage = "Rendering the blocks failed." 'A message that properly informs the user and the devs (silent errors will be logged nonetheless)
      .SetCallArgs "No args" 'If the sub takes args put the here like ("sExample:=" & sExample, "lExample:=" & lExample)
   End With
'Fixed, don't change
Try: On Error GoTo Catch


'>>>>>>> Your code here

   Dim lOutRow As Long
   Dim lOutCol As Long
   Dim lBlockIndex As Long
   Dim uBlockSpec As u_f_BlockSpecRenderer
   Dim oDictContext As Scripting.Dictionary
   Dim lBlockRowHeight As Long
    
   lOutRow = lStartRowOut
   lOutCol = lStartColOut
   
   For lBlockIndex = LBound(uaBlockSpecs) To UBound(uaBlockSpecs)
      uBlockSpec = uaBlockSpecs(lBlockIndex)
      If oDictBlocksData.Exists(uBlockSpec.sBlockKey) Then
         Set oDictContext = oDictBlocksData(uBlockSpec.sBlockKey)
      Else
         Set oDictContext = oDict_f_p_BuildEmptyContext()
      End If
      lBlockRowHeight = l_m_RenderOneBlock(oWksOut, uBlockSpec, oDictContext, lOutRow, lOutCol)
      lOutRow = lOutRow + lBlockRowHeight
   Next lBlockIndex

'End of your code <<<<<<<


'Fixed, don't change
Finally: On Error Resume Next


'>>>>>>> Your code here
   lNextStartRow = lOutRow


'End of your code <<<<<<<


'>>>>>>> Your custom settings here
   If oC_Me.oC_prop_r_Error Is Nothing Then b_f_p_RenderBlocks = True 'reports execution as successful to caller
'Fixed, don't change
   Exit Function
HandleError: af_pM_ErrorHandling.af_p_Hook_ErrorHandling_LowerLevel


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
End Function


Private Function l_m_RenderOneBlock( _
   ByRef oWksOut As Worksheet, _
   ByRef uBlockSpec As u_f_BlockSpecRenderer, _
   ByVal oDictContext As Scripting.Dictionary, _
   ByVal lOutTopRow As Long, _
   ByVal lOutLeftColumn As Long _
) As Long
    
   Dim lMaxBottomRow As Long
   Dim lIndex As Long
   Dim lBottomRowFixLane As Long
   Dim lBottomRowRepeatableLane As Long
   Dim uaRelsLaneSpec() As u_f_LaneSpecRenderer
   Dim lBottomRowRelativeLanes As Long
   
   mValidateRepeaters uBlockSpec, oDictContext
    
   lMaxBottomRow = 0
    
   ' 1) FIX zuerst (Template-Position)
   If b_m_HasElementsLanes(uBlockSpec.uaLanes) Then
      For lIndex = LBound(uBlockSpec.uaLanes) To UBound(uBlockSpec.uaLanes)
         If UCase$(uBlockSpec.uaLanes(lIndex).sLaneType) = s_m_LANE_TAG_FIX Then
            lBottomRowFixLane = l_m_WriteFixLane(oWksOut, uBlockSpec, uBlockSpec.uaLanes(lIndex), oDictContext, lOutTopRow, lOutLeftColumn)
            If lBottomRowFixLane > lMaxBottomRow Then lMaxBottomRow = lBottomRowFixLane
         End If
      Next lIndex
   End If
    
   ' 2) REP danach (expandiert)
   If b_m_HasElementsLanes(uBlockSpec.uaLanes) Then
      For lIndex = LBound(uBlockSpec.uaLanes) To UBound(uBlockSpec.uaLanes)
         If UCase$(uBlockSpec.uaLanes(lIndex).sLaneType) = s_m_LANE_TAG_REPEATABLE Then
            lBottomRowRepeatableLane = l_m_WriteRepLane(oWksOut, uBlockSpec, uBlockSpec.uaLanes(lIndex), oDictContext, lOutTopRow, lOutLeftColumn)
            If lBottomRowRepeatableLane > lMaxBottomRow Then lMaxBottomRow = lBottomRowRepeatableLane
         End If
      Next lIndex
   End If
    
   ' 3) REL zum Schluss (verschieben, wenn nötig)
   uaRelsLaneSpec = ua_m_SortRelLanesByTopRel(uBlockSpec.uaLanes)
   If b_m_HasElementsLanes(uaRelsLaneSpec) Then
      For lIndex = LBound(uaRelsLaneSpec) To UBound(uaRelsLaneSpec)
         lBottomRowRelativeLanes = l_m_WriteRelLane(oWksOut, uBlockSpec, uaRelsLaneSpec(lIndex), oDictContext, lOutTopRow, lOutLeftColumn, lMaxBottomRow)
         If lBottomRowRelativeLanes > lMaxBottomRow Then lMaxBottomRow = lBottomRowRelativeLanes
      Next lIndex
   End If
    
   If lMaxBottomRow = 0 Then lMaxBottomRow = uBlockSpec.lRowsCount
   l_m_RenderOneBlock = lMaxBottomRow
End Function

Private Function ua_m_SortRelLanesByTopRel(ByRef uaLaneSpecs() As u_f_LaneSpecRenderer) As u_f_LaneSpecRenderer()
   Dim uaLaneSpecsBuffer() As u_f_LaneSpecRenderer
   Dim lIndex As Long, lRelLaneCount As Long
   Dim uTempLaneSpec As u_f_LaneSpecRenderer
   Dim bSwapped As Boolean
   
   lRelLaneCount = -1
   If b_m_HasElementsLanes(uaLaneSpecs) Then
      For lIndex = LBound(uaLaneSpecs) To UBound(uaLaneSpecs)
         If UCase$(uaLaneSpecs(lIndex).sLaneType) = s_m_LANE_TAG_RELATIVE Then
            lRelLaneCount = lRelLaneCount + 1
            ReDim Preserve uaLaneSpecsBuffer(0 To lRelLaneCount)
            uaLaneSpecsBuffer(lRelLaneCount) = uaLaneSpecs(lIndex)
         End If
      Next lIndex
   End If
   If lRelLaneCount < 0 Then
      ua_m_SortRelLanesByTopRel = uaLaneSpecsBuffer  ' uaLaneSpecsBuffer ist uninitialisiert ? ok
      Exit Function
   End If

   
   ' sortieren nach TopRel
   If b_m_HasElementsLanes(uaLaneSpecsBuffer) Then
      Do
         bSwapped = False
         For lIndex = LBound(uaLaneSpecsBuffer) To UBound(uaLaneSpecsBuffer) - 1
            If uaLaneSpecsBuffer(lIndex).lTopRel > uaLaneSpecsBuffer(lIndex + 1).lTopRel Then
               uTempLaneSpec = uaLaneSpecsBuffer(lIndex): uaLaneSpecsBuffer(lIndex) = uaLaneSpecsBuffer(lIndex + 1): uaLaneSpecsBuffer(lIndex + 1) = uTempLaneSpec
               bSwapped = True
            End If
         Next lIndex
      Loop While bSwapped
   End If
   ua_m_SortRelLanesByTopRel = uaLaneSpecsBuffer

End Function

Private Function b_m_HasElementsLanes(ByRef uLaneSpecs() As u_f_LaneSpecRenderer) As Boolean
   On Error Resume Next
   b_m_HasElementsLanes = (UBound(uLaneSpecs) >= LBound(uLaneSpecs))
   If Err.Number <> 0 Then
      b_m_HasElementsLanes = False
      Err.Clear
   End If
End Function

Private Function l_m_WriteFixLane( _
   ByRef oWksOut As Worksheet, _
   ByRef uBlockSpec As u_f_BlockSpecRenderer, _
   ByRef uLaneSpec As u_f_LaneSpecRenderer, _
   ByVal oDictContext As Scripting.Dictionary, _
   ByVal lOutTopRow As Long, _
   ByVal lOutLeftCol As Long _
) As Long
    
   Dim lRow As Long, lCol As Long
   Dim vaValues() As Variant
   Dim saTokens() As String
   Dim uCellSpec As u_f_CellSpecRenderer
   Dim lBottomRelRow As Long
   
   For lRow = 1 To uLaneSpec.lRowsCount
      ReDim vaValues(1 To 1, 1 To uLaneSpec.lColsCount)
      ReDim saTokens(1 To uLaneSpec.lColsCount)
      For lCol = 1 To uLaneSpec.lColsCount
         uCellSpec = u_m_GetCellSpecFromLane(uLaneSpec, lRow, lCol)
         vaValues(1, lCol) = s_m_ReplaceAllTemplateTags(uCellSpec.sTemplateText, uCellSpec.sPlaceholderList, oDictContext, Nothing)
         saTokens(lCol) = uCellSpec.sStyleToken
      Next lCol
      With oWksOut.Range( _
      oWksOut.Cells(lOutTopRow + uLaneSpec.lTopRel + lRow - 1, lOutLeftCol + uLaneSpec.lLeftRel), _
      oWksOut.Cells(lOutTopRow + uLaneSpec.lTopRel + lRow - 1, lOutLeftCol + uLaneSpec.lLeftRel + uLaneSpec.lColsCount - 1))
         .Value = vaValues
         mApplyStylesRow .Cells, saTokens
      End With
   Next lRow
   lBottomRelRow = uLaneSpec.lTopRel + uLaneSpec.lRowsCount - 1
   If uLaneSpec.lPadAfterRows > 0 Then
      lBottomRelRow = l_m_WritePadRows(oWksOut, lOutTopRow, lOutLeftCol, uLaneSpec, lBottomRelRow)
   End If
   l_m_WriteFixLane = lBottomRelRow
End Function

Private Function l_m_WriteRepLane( _
   ByRef oWksOut As Worksheet, _
   ByRef uBlockSpec As u_f_BlockSpecRenderer, _
   ByRef uLaneSpec As u_f_LaneSpecRenderer, _
   ByVal oDictContext As Scripting.Dictionary, _
   ByVal lOutTopRow As Long, _
   ByVal lOutLeftCol As Long) As Long
   
   Dim oColItems As Collection
   Dim lItemCount As Long
   Dim lIndex As Long, lRow As Long, lCol As Long
   Dim vaValues() As Variant
   Dim saTokens() As String
   Dim uCellSpec As u_f_CellSpecRenderer
   Dim lBottomRelRow As Long
   
   Set oColItems = oCol_m_ResolveRepCollection(oDictContext, uLaneSpec.sKey)
   
   If oColItems Is Nothing Then
      lItemCount = 0
   Else
      lItemCount = oColItems.Count
   End If
    
   For lIndex = 1 To lItemCount
      For lRow = 1 To uLaneSpec.lRowsCount
         ReDim vaValues(1 To 1, 1 To uLaneSpec.lColsCount)
         ReDim saTokens(1 To uLaneSpec.lColsCount)
         For lCol = 1 To uLaneSpec.lColsCount
            uCellSpec = u_m_GetCellSpecFromLane(uLaneSpec, lRow, lCol)
            vaValues(1, lCol) = s_m_ReplaceAllTemplateTags(uCellSpec.sTemplateText, uCellSpec.sPlaceholderList, oDictContext, oColItems(lIndex))
            saTokens(lCol) = uCellSpec.sStyleToken
         Next lCol
         With oWksOut.Range( _
         oWksOut.Cells(lOutTopRow + uLaneSpec.lTopRel + (lIndex - 1) * uLaneSpec.lRowsCount + lRow - 1, lOutLeftCol + uLaneSpec.lLeftRel), _
         oWksOut.Cells(lOutTopRow + uLaneSpec.lTopRel + (lIndex - 1) * uLaneSpec.lRowsCount + lRow - 1, lOutLeftCol + uLaneSpec.lLeftRel + uLaneSpec.lColsCount - 1))
            .Value = vaValues
            mApplyStylesRow .Cells, saTokens
         End With
      Next lRow
   Next lIndex
    
   lBottomRelRow = uLaneSpec.lTopRel + Application.Max(lItemCount * uLaneSpec.lRowsCount, uLaneSpec.lRowsCount) - 1
   If uLaneSpec.lPadAfterRows > 0 Then
      lBottomRelRow = l_m_WritePadRows(oWksOut, lOutTopRow, lOutLeftCol, uLaneSpec, lBottomRelRow)
   End If
   l_m_WriteRepLane = lBottomRelRow

End Function

Private Function l_m_WriteRelLane( _
   ByRef oWksOut As Worksheet, _
   ByRef uBlockSpec As u_f_BlockSpecRenderer, _
   ByRef uLaneSpec As u_f_LaneSpecRenderer, _
   ByVal oDictContext As Scripting.Dictionary, _
   ByVal lOutTopRow As Long, _
   ByVal lOutLeftCol As Long, _
   ByVal lCurrentMaxBottomRow As Long _
) As Long
    
   Dim lEffectiveTopRel As Long
   Dim vaValues() As Variant
   Dim lRow As Long, lCol As Long
   Dim saTokens() As String
   Dim uCellSpec As u_f_CellSpecRenderer
   Dim lBottomRelRow As Long
   
   lEffectiveTopRel = uLaneSpec.lTopRel
   If lCurrentMaxBottomRow >= lEffectiveTopRel Then
      lEffectiveTopRel = lCurrentMaxBottomRow + 1
   End If
   
   For lRow = 1 To uLaneSpec.lRowsCount
      ReDim vaValues(1 To 1, 1 To uLaneSpec.lColsCount)
      ReDim saTokens(1 To uLaneSpec.lColsCount)
      For lCol = 1 To uLaneSpec.lColsCount
         uCellSpec = u_m_GetCellSpecFromLane(uLaneSpec, lRow, lCol)
         vaValues(1, lCol) = s_m_ReplaceAllTemplateTags(uCellSpec.sTemplateText, uCellSpec.sPlaceholderList, oDictContext, Nothing)
         saTokens(lCol) = uCellSpec.sStyleToken
      Next lCol
      With oWksOut.Range( _
         oWksOut.Cells(lOutTopRow + lEffectiveTopRel + lRow - 1, lOutLeftCol + uLaneSpec.lLeftRel), _
         oWksOut.Cells(lOutTopRow + lEffectiveTopRel + lRow - 1, lOutLeftCol + uLaneSpec.lLeftRel + uLaneSpec.lColsCount - 1))
         .Value = vaValues
         mApplyStylesRow .Cells, saTokens
      End With
   Next lRow
   
   lBottomRelRow = lEffectiveTopRel + uLaneSpec.lRowsCount - 1
   If uLaneSpec.lPadAfterRows > 0 Then
      lBottomRelRow = l_m_WritePadRowsAt(oWksOut, lOutTopRow, lOutLeftCol, uLaneSpec, lBottomRelRow + 1, uLaneSpec.lLeftRel)
   End If
   l_m_WriteRelLane = lBottomRelRow

End Function

Private Function l_m_WritePadRows( _
   ByRef oWksOut As Worksheet, _
   ByVal lOutTopRow As Long, _
   ByVal lOutLeftCol As Long, _
   ByRef uLaneSpec As u_f_LaneSpecRenderer, _
   ByVal lCurrentBottomRowRel As Long _
) As Long
    
    Dim lNextRowRel As Long
    lNextRowRel = lCurrentBottomRowRel + 1
    
    l_m_WritePadRows = l_m_WritePadRowsAt(oWksOut, lOutTopRow, lOutLeftCol, uLaneSpec, lNextRowRel, uLaneSpec.lLeftRel)
End Function

Private Function l_m_WritePadRowsAt( _
   ByRef oWksOut As Worksheet, _
   ByVal lOutTopRow As Long, _
   ByVal lOutLeftCol As Long, _
   ByRef uLaneSpec As u_f_LaneSpecRenderer, _
   ByVal lStartRowRel As Long, _
   ByVal lStartColRel As Long _
) As Long
    
   Dim i As Long
   Dim vaValues() As Variant
   Dim saTokens() As String
    
   For i = 1 To uLaneSpec.lPadAfterRows
      ReDim vaValues(1 To 1, 1 To uLaneSpec.lColsCount)
      ReDim saTokens(1 To uLaneSpec.lColsCount)
      With oWksOut.Range( _
         oWksOut.Cells(lOutTopRow + lStartRowRel + i - 1, lOutLeftCol + lStartColRel), _
         oWksOut.Cells(lOutTopRow + lStartRowRel + i - 1, lOutLeftCol + lStartColRel + uLaneSpec.lColsCount - 1) _
      )
         .Value2 = vaValues
         mApplyStylesRow .Cells, saTokens
      End With
   Next i
   l_m_WritePadRowsAt = lStartRowRel + uLaneSpec.lPadAfterRows - 1

End Function

Private Function u_m_GetCellSpecFromLane( _
   ByRef uLaneSpec As u_f_LaneSpecRenderer, _
   ByVal lRelRow As Long, _
   ByVal lRelCol As Long _
) As u_f_CellSpecRenderer
   
   Dim lIndex As Long
   lIndex = (lRelRow - 1) * uLaneSpec.lColsCount + lRelCol
   u_m_GetCellSpecFromLane = uLaneSpec.uaCells(lIndex)

End Function

Private Sub mApplyStylesRow(ByRef oRngRow As Range, ByRef saStyleTokens() As String)
   Dim i As Long
   Dim sToken As String
   For i = 1 To oRngRow.Columns.Count
      sToken = saStyleTokens(i)
      If Len(sToken) > 0 Then
         On Error Resume Next
         oRngRow.Cells(1, i).Style = sToken
         On Error GoTo 0
         ApplyBordersForToken oRngRow.Cells(1, i), sToken
      End If
   Next i
End Sub


' Purpose: replaces template tags with actual values
Private Function s_m_ReplaceAllTemplateTags( _
   ByVal sTemplateText As String, _
   ByVal sListKeys As String, _
   ByVal oDictContext As Scripting.Dictionary, _
   ByVal vItem As Variant _
) As String
    
   Dim sOutput As String
   Dim saKeys() As String
   Dim i As Long
   Dim sKey As String
   Dim sValue As String
   
   sOutput = sTemplateText
   If Len(sListKeys) = 0 Then
      s_m_ReplaceAllTemplateTags = sOutput
      Exit Function
   End If
   
   saKeys = Split(sListKeys, s_m_SEPARATOR)
   
   For i = LBound(saKeys) To UBound(saKeys)
      sKey = saKeys(i)
      sValue = s_m_ResolveValue(sKey, oDictContext, vItem)
      sOutput = Replace(sOutput, "{{" & sKey & "}}", sValue)
   Next i
   
   s_m_ReplaceAllTemplateTags = sOutput
   
End Function

' Purpose: replaces value tag with actual value
Private Function s_m_ResolveValue(ByVal sKey As String, ByVal oDictContext As Scripting.Dictionary, ByVal vItem As Variant) As String
   On Error Resume Next
   If InStr(1, sKey, "[i].", vbTextCompare) > 0 Then
      If Not IsEmpty(vItem) Then
         If vItem.Exists(sKey) Then
            s_m_ResolveValue = CStr(vItem(sKey))
         Else
            s_m_ResolveValue = ""
         End If
      Else
         s_m_ResolveValue = ""
      End If
   ElseIf Left$(sKey, 7) = "Totals." Then
      If oDictContext(s_m_KEY_TOTALS).Exists(sKey) Then
         s_m_ResolveValue = CStr(oDictContext(s_m_KEY_TOTALS)(sKey))
      Else
         s_m_ResolveValue = ""
      End If
   Else
       If oDictContext(s_m_KEY_HEADER).Exists(sKey) Then
         s_m_ResolveValue = CStr(oDictContext(s_m_KEY_HEADER)(sKey))
      Else
         s_m_ResolveValue = ""
      End If
   End If
   On Error GoTo 0
End Function

' Purpose: gets the collection with repeatable items for the provided lane key
Private Function oCol_m_ResolveRepCollection(ByVal oDictContext As Scripting.Dictionary, ByVal sLaneKey As String) As Collection
    On Error Resume Next
    Dim oDictRepeaters As Scripting.Dictionary
    Set oDictRepeaters = oDictContext(s_m_KEY_REPEATERS)
    If Not oDictRepeaters Is Nothing Then
        If oDictRepeaters.Exists(sLaneKey) Then Set oCol_m_ResolveRepCollection = oDictRepeaters(sLaneKey)
    End If
    On Error GoTo 0
End Function

Private Sub mValidateRepeaters(ByRef uBlockSpec As u_f_BlockSpecRenderer, ByVal oDictBlockData As Scripting.Dictionary)
   Dim lIndex As Long
' TODO: [+] Refactor error handling to framework logic
   For lIndex = LBound(uBlockSpec.uaLanes) To UBound(uBlockSpec.uaLanes)
      If UCase$(uBlockSpec.uaLanes(lIndex).sLaneType) = s_m_LANE_TAG_REPEATABLE Then
         If Not oDictBlockData(s_m_KEY_REPEATERS).Exists(uBlockSpec.uaLanes(lIndex).sKey) Then
            Err.Raise 5, , "Fehlende Repeater-Daten: blk_" & uBlockSpec.sBlockKey & _
                           " ? rep key '" & uBlockSpec.uaLanes(lIndex).sKey & "'"
         End If
      End If
   Next lIndex
End Sub




