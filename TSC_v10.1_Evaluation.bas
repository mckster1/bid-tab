Attribute VB_Name = "TSC_v10_Evaluation"
Option Explicit

' Private color constants — apply only to scope cell highlighting in this module
Private Const CLR_EXCLUDED As Long = 11842815    ' RGB(255,180,180) – non-anomaly EXCLUDED cell
Private Const CLR_SPELLCHECK As Long = 15128749  ' RGB(173,216,230) – description with possible misspelling

'===========================
' PUBLIC ORCHESTRATOR (button-assignable)
'===========================
' Refreshes the evaluation summary (averages, lowest bidder) and all highlights
' on the active trade tab. Safe to call from buttons or from Workbook_SheetChange.
'
' AUTO-CHANGE SETUP (one-time, paste into ThisWorkbook module):
'
'   Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
'       If StrComp(Sh.Name, SHEET_TEMPLATE, vbTextCompare) = 0 Then Exit Sub
'       If StrComp(Sh.Name, SHEET_CONFIG, vbTextCompare) = 0 Then Exit Sub
'       If Left$(Sh.Name, 3) = "zz_" Then Exit Sub
'       If Not Application.EnableEvents Then Exit Sub
'       On Error Resume Next
'       Application.EnableEvents = False
'       RefreshHighlights_v10
'       Application.EnableEvents = True
'       On Error GoTo 0
'   End Sub
'
Public Sub RefreshHighlights_v10()
    Dim ws As Worksheet: Set ws = ActiveSheet

    ' Guard: only run on live trade tabs
    If StrComp(ws.Name, SHEET_TEMPLATE, vbTextCompare) = 0 Then Exit Sub
    If StrComp(ws.Name, SHEET_CONFIG, vbTextCompare) = 0 Then Exit Sub
    If Left$(ws.Name, 3) = "zz_" Then Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    On Error GoTo CleanUp
    UpdateEvalSummary ws
    HighlightLowestBid ws
    HighlightAnomalies ws

CleanUp:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

'===========================
' FORMAT AND VALIDATE (button-assignable)
' Auto-corrects I/E/U abbreviations in scope cells, applies color coding,
' and highlights description cells with possible spelling issues.
'===========================
Public Sub FormatAndValidate_v10()
    Dim ws As Worksheet: Set ws = ActiveSheet

    If StrComp(ws.Name, SHEET_TEMPLATE, vbTextCompare) = 0 Then Exit Sub
    If StrComp(ws.Name, SHEET_CONFIG, vbTextCompare) = 0 Then Exit Sub
    If Left$(ws.Name, 3) = "zz_" Then Exit Sub

    Dim anchor As Long: anchor = FindRowKeyInColB(ws, KEY_ADD_SCOPE)
    If anchor = 0 Then MsgBox "Can't find '" & KEY_ADD_SCOPE & "' in col B.", vbExclamation: Exit Sub

    Dim lastCol As Long
    lastCol = ws.Cells(ROW_WIZ_ACTION, ws.Columns.Count).End(xlToLeft).Column

    Dim bidCols() As Long, nBid As Long
    ReDim bidCols(1 To lastCol - COL_BIDDER_START + 1)
    Dim c As Long
    For c = COL_BIDDER_START To lastCol
        If Not IsBidderColEmpty(ws, c) Then
            nBid = nBid + 1
            bidCols(nBid) = c
        End If
    Next c
    If nBid = 0 Then MsgBox "No bidder columns found.", vbInformation: Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Dim r As Long, i As Long, fixCount As Long, spellCount As Long
    Dim desc As String

    ' Process scope rows
    For r = ROW_SCOPE_START To anchor - 1
        desc = Trim$(CStr(ws.Cells(r, COL_DESC).Value))
        If Len(desc) = 0 Then GoTo NextFVScope
        If Left$(UCase$(desc), 10) = "EXCEPTION:" Then GoTo NextFVScope

        ' Spell-check description
        If HasSpellingIssue(desc) Then
            If ws.Cells(r, COL_DESC).Interior.Color <> CLR_SPELLCHECK Then
                ws.Cells(r, COL_DESC).Interior.Color = CLR_SPELLCHECK
                spellCount = spellCount + 1
            End If
        Else
            If ws.Cells(r, COL_DESC).Interior.Color = CLR_SPELLCHECK Then
                ws.Cells(r, COL_DESC).Interior.ColorIndex = xlNone
            End If
        End If

        For i = 1 To nBid
            fixCount = fixCount + ApplyFVCell(ws, r, bidCols(i))
        Next i
NextFVScope:
    Next r

    ' Process alternate rows (below ADD ALTERNATE anchor)
    Dim altAnchor As Long: altAnchor = FindRowKeyInColB(ws, KEY_ADD_ALT)
    If altAnchor > 0 Then
        Dim altDataLast As Long: altDataLast = ws.Cells(ws.Rows.Count, COL_DESC).End(xlUp).Row
        For r = altAnchor + 1 To altDataLast
            desc = Trim$(CStr(ws.Cells(r, COL_DESC).Value))
            If Len(desc) = 0 Then GoTo NextFVAlt
            If InStr(UCase$(desc), "ADJUSTED BASE BID") > 0 Then GoTo NextFVAlt
            If UCase$(Left$(desc, 9)) <> "ALTERNATE" Then GoTo NextFVAlt
            For i = 1 To nBid
                fixCount = fixCount + ApplyFVCell(ws, r, bidCols(i))
            Next i
NextFVAlt:
        Next r
    End If

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    ' Run highlights to apply anomaly coloring on top of corrected values
    RefreshHighlights_v10

    Dim summMsg As String
    summMsg = "Format & Validate complete." & vbCrLf
    If fixCount > 0 Then summMsg = summMsg & "  " & fixCount & " cell(s) auto-corrected." & vbCrLf
    If spellCount > 0 Then summMsg = summMsg & "  " & spellCount & " description(s) flagged for spelling review (light blue)." & vbCrLf
    If fixCount = 0 And spellCount = 0 Then summMsg = summMsg & "  No issues found."
    MsgBox summMsg, vbInformation, "Format & Validate"
End Sub

' Applies auto-correct and color coding to a single bidder cell.
' Returns 1 if a text correction was made, 0 otherwise.
Private Function ApplyFVCell(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long) As Long
    If ws.Cells(r, c).HasFormula Then Exit Function   ' never overwrite formulas

    Dim cv As Variant: cv = ws.Cells(r, c).Value
    Dim s As String: s = Trim$(CStr(cv))
    Dim su As String: su = UCase$(s)

    Select Case su
        Case "I", "INC", "INCL"
            ws.Cells(r, c).Value = TXT_INCLUDED
            ws.Cells(r, c).Interior.ColorIndex = xlNone
            ApplyFVCell = 1
        Case "E", "EXC", "EXCL"
            ws.Cells(r, c).Value = TXT_EXCLUDED
            ws.Cells(r, c).Interior.Color = CLR_EXCLUDED
            ApplyFVCell = 1
        Case "U", "UNCONF"
            ws.Cells(r, c).Value = TXT_UNCONF
            ws.Cells(r, c).Interior.Color = RGB_LIGHT_YELLOW
            ApplyFVCell = 1
        Case TXT_INCLUDED
            ws.Cells(r, c).Interior.ColorIndex = xlNone
        Case TXT_EXCLUDED
            ws.Cells(r, c).Interior.Color = CLR_EXCLUDED
        Case TXT_UNCONF
            ws.Cells(r, c).Interior.Color = RGB_LIGHT_YELLOW
        Case ""
            ' empty — leave as-is
        Case Else
            If IsNumeric(cv) Then
                ws.Cells(r, c).Interior.ColorIndex = xlNone
            ElseIf Len(s) > 0 Then
                ' Non-keyword, non-numeric text — flag as suspicious
                ws.Cells(r, c).Interior.Color = RGB_ORANGE
            End If
    End Select
End Function

'===========================
' EVALUATION SUMMARY
' Writes averages and lowest bidder name to the eval rows in col G (labels) and col H (values)
'===========================
Private Sub UpdateEvalSummary(ByVal ws As Worksheet)
    ws.Cells(ROW_LOWEST_BIDDER, COL_HDR).Value = TXT_LBL_LOWEST
    ws.Cells(ROW_AVG_BASE, COL_HDR).Value = TXT_LBL_AVG_BASE
    ws.Cells(ROW_AVG_ADJ, COL_HDR).Value = TXT_LBL_AVG_ADJ

    Dim lastCol As Long
    lastCol = ws.Cells(ROW_WIZ_ACTION, ws.Columns.Count).End(xlToLeft).Column

    Dim sumBase As Double, sumAdj As Double
    Dim cntBase As Long, cntAdj As Long
    Dim lowestAdj As Double, lowestCol As Long, hasAdj As Boolean

    Dim c As Long
    For c = COL_BIDDER_START To lastCol
        If IsBidderColEmpty(ws, c) Then GoTo NextBidder
        If UCase$(Trim$(CStr(ws.Cells(ROW_WIZ_ACTION, c).Value))) = BIDDER_EXCLUDE Then GoTo NextBidder

        Dim vBase As Variant: vBase = ws.Cells(ROW_BASE_BID, c).Value
        If IsNumeric(vBase) And Len(Trim$(CStr(vBase))) > 0 Then
            sumBase = sumBase + CDbl(vBase)
            cntBase = cntBase + 1
        End If

        Dim vAdj As Variant: vAdj = ws.Cells(ROW_ADJ_BASE, c).Value
        If IsNumeric(vAdj) And Len(Trim$(CStr(vAdj))) > 0 Then
            Dim dAdj As Double: dAdj = CDbl(vAdj)
            sumAdj = sumAdj + dAdj
            cntAdj = cntAdj + 1
            If Not hasAdj Or dAdj < lowestAdj Then
                lowestAdj = dAdj
                lowestCol = c
                hasAdj = True
            End If
        End If

NextBidder:
    Next c

    If cntBase > 0 Then
        ws.Cells(ROW_AVG_BASE, COL_BUDGET).Value = sumBase / cntBase
        ws.Cells(ROW_AVG_BASE, COL_BUDGET).NumberFormat = "$#,##0"
    Else
        ws.Cells(ROW_AVG_BASE, COL_BUDGET).ClearContents
    End If

    If cntAdj > 0 Then
        ws.Cells(ROW_AVG_ADJ, COL_BUDGET).Value = sumAdj / cntAdj
        ws.Cells(ROW_AVG_ADJ, COL_BUDGET).NumberFormat = "$#,##0"
    Else
        ws.Cells(ROW_AVG_ADJ, COL_BUDGET).ClearContents
    End If

    If lowestCol > 0 Then
        Dim bidName As String: bidName = Trim$(CStr(ws.Cells(2, lowestCol).Value))
        If Len(bidName) = 0 Then bidName = ColLetter(lowestCol)
        ws.Cells(ROW_LOWEST_BIDDER, COL_BUDGET).Value = bidName
    Else
        ws.Cells(ROW_LOWEST_BIDDER, COL_BUDGET).ClearContents
    End If
End Sub

'===========================
' LOWEST BID HIGHLIGHT
'===========================
Private Sub HighlightLowestBid(ByVal ws As Worksheet)
    Dim lastCol As Long
    lastCol = ws.Cells(ROW_WIZ_ACTION, ws.Columns.Count).End(xlToLeft).Column

    Dim c As Long, r As Long
    For c = COL_BIDDER_START To lastCol
        If Not IsBidderColEmpty(ws, c) Then
            For r = 2 To ROW_WIZ_ACTION
                If ws.Cells(r, c).Interior.Color = RGB_LIGHT_GREEN Then
                    ws.Cells(r, c).Interior.ColorIndex = xlNone
                End If
            Next r
        End If
    Next c

    Dim lowestAdj As Double, lowestCol As Long, hasAdj As Boolean
    For c = COL_BIDDER_START To lastCol
        If IsBidderColEmpty(ws, c) Then GoTo NextCol
        If UCase$(Trim$(CStr(ws.Cells(ROW_WIZ_ACTION, c).Value))) = BIDDER_EXCLUDE Then GoTo NextCol

        Dim v As Variant: v = ws.Cells(ROW_ADJ_BASE, c).Value
        If IsNumeric(v) And Len(Trim$(CStr(v))) > 0 Then
            Dim dv As Double: dv = CDbl(v)
            If Not hasAdj Or dv < lowestAdj Then
                lowestAdj = dv
                lowestCol = c
                hasAdj = True
            End If
        End If
NextCol:
    Next c

    If lowestCol > 0 Then
        For r = 2 To ROW_WIZ_ACTION
            ws.Cells(r, lowestCol).Interior.Color = RGB_LIGHT_GREEN
        Next r
    End If
End Sub

'===========================
' ANOMALY HIGHLIGHTING
' - UNCONFIRMED → light yellow
' - EXCLUDED, ≥60% of others INCLUDED → orange (anomaly)
' - EXCLUDED, below threshold → light red (CLR_EXCLUDED)
' - Skips exception/exclusion rows
'===========================
Private Sub HighlightAnomalies(ByVal ws As Worksheet)
    Dim anchor As Long: anchor = FindRowKeyInColB(ws, KEY_ADD_SCOPE)
    If anchor = 0 Then Exit Sub

    Dim lastCol As Long
    lastCol = ws.Cells(ROW_WIZ_ACTION, ws.Columns.Count).End(xlToLeft).Column

    Dim bidCols() As Long, nBid As Long
    ReDim bidCols(1 To lastCol - COL_BIDDER_START + 1)
    Dim c As Long
    For c = COL_BIDDER_START To lastCol
        If Not IsBidderColEmpty(ws, c) Then
            nBid = nBid + 1
            bidCols(nBid) = c
        End If
    Next c
    If nBid = 0 Then Exit Sub

    Dim r As Long, i As Long
    For r = ROW_SCOPE_START To anchor - 1
        Dim desc As String: desc = Trim$(CStr(ws.Cells(r, COL_DESC).Value))
        If Len(desc) = 0 Then GoTo NextRow
        If Left$(UCase$(desc), 10) = "EXCEPTION:" Then GoTo NextRow

        Dim incCount As Long: incCount = 0
        For i = 1 To nBid
            Dim cv As String: cv = UCase$(Trim$(CStr(ws.Cells(r, bidCols(i)).Value)))
            If cv = TXT_INCLUDED Or IsNumericResponse(cv) Then incCount = incCount + 1
        Next i

        For i = 1 To nBid
            Dim cellVal As String: cellVal = UCase$(Trim$(CStr(ws.Cells(r, bidCols(i)).Value)))

            If cellVal = TXT_UNCONF Then
                ws.Cells(r, bidCols(i)).Interior.Color = RGB_LIGHT_YELLOW

            ElseIf cellVal = TXT_EXCLUDED Then
                If nBid >= 2 And incCount / nBid >= 0.6 Then
                    ws.Cells(r, bidCols(i)).Interior.Color = RGB_ORANGE
                Else
                    ws.Cells(r, bidCols(i)).Interior.Color = CLR_EXCLUDED
                End If

            Else
                ' Clear stale highlights from cells that now have real data
                Dim curColor As Long: curColor = ws.Cells(r, bidCols(i)).Interior.Color
                If curColor = RGB_LIGHT_YELLOW Or curColor = RGB_ORANGE Or curColor = CLR_EXCLUDED Then
                    ws.Cells(r, bidCols(i)).Interior.ColorIndex = xlNone
                End If
            End If
        Next i

NextRow:
    Next r
End Sub

Private Function IsNumericResponse(ByVal s As String) As Boolean
    IsNumericResponse = IsNumeric(s) And Len(s) > 0
End Function

'===========================
' SPELL CHECK HELPERS
'===========================
Private Function HasSpellingIssue(ByVal desc As String) As Boolean
    On Error GoTo HasSpellingExit
    If Len(desc) <= 3 Then Exit Function
    If desc = UCase$(desc) Then Exit Function   ' all-caps: code/acronym
    If desc Like "#*" Then Exit Function         ' starts with digit

    Dim words() As String: words = Split(desc, " ")
    Dim w As Long, wd As String
    For w = 0 To UBound(words)
        wd = StripPunctuation(Trim$(words(w)))
        If Len(wd) < 3 Then GoTo NextWord
        If wd = UCase$(wd) Then GoTo NextWord
        If wd Like "#*" Then GoTo NextWord
        If Not Application.CheckSpelling(wd) Then
            HasSpellingIssue = True
            Exit Function
        End If
NextWord:
    Next w
    Exit Function
HasSpellingExit:
    HasSpellingIssue = False
End Function

Private Function StripPunctuation(ByVal w As String) As String
    Dim i As Long, ch As String, out As String
    For i = 1 To Len(w)
        ch = Mid$(w, i, 1)
        If ch Like "[A-Za-z'-]" Then out = out & ch
    Next i
    StripPunctuation = out
End Function
