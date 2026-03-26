Attribute VB_Name = "TSC_v10_Evaluation"
Option Explicit

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
' EVALUATION SUMMARY
' Writes averages and lowest bidder name to the eval rows in col G (labels) and col H (values)
'===========================
Private Sub UpdateEvalSummary(ByVal ws As Worksheet)
    ' Write / refresh labels
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

        ' Base Bid
        Dim vBase As Variant: vBase = ws.Cells(ROW_BASE_BID, c).Value
        If IsNumeric(vBase) And Len(Trim$(CStr(vBase))) > 0 Then
            sumBase = sumBase + CDbl(vBase)
            cntBase = cntBase + 1
        End If

        ' Adjusted Base Bid
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

    ' Write average base bid
    If cntBase > 0 Then
        ws.Cells(ROW_AVG_BASE, COL_BUDGET).Value = sumBase / cntBase
        ws.Cells(ROW_AVG_BASE, COL_BUDGET).NumberFormat = "$#,##0"
    Else
        ws.Cells(ROW_AVG_BASE, COL_BUDGET).ClearContents
    End If

    ' Write average adjusted base bid
    If cntAdj > 0 Then
        ws.Cells(ROW_AVG_ADJ, COL_BUDGET).Value = sumAdj / cntAdj
        ws.Cells(ROW_AVG_ADJ, COL_BUDGET).NumberFormat = "$#,##0"
    Else
        ws.Cells(ROW_AVG_ADJ, COL_BUDGET).ClearContents
    End If

    ' Write lowest bidder name
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
' Highlights the lowest adjusted base bid bidder column header rows (2–8) in green.
' Clears previous green highlights before applying new ones.
'===========================
Private Sub HighlightLowestBid(ByVal ws As Worksheet)
    Dim lastCol As Long
    lastCol = ws.Cells(ROW_WIZ_ACTION, ws.Columns.Count).End(xlToLeft).Column

    ' Clear existing green highlights from all bidder header rows
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

    ' Find lowest adjusted base bid among included bidders
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

    ' Apply green highlight to the winning column's header rows
    If lowestCol > 0 Then
        For r = 2 To ROW_WIZ_ACTION
            ws.Cells(r, lowestCol).Interior.Color = RGB_LIGHT_GREEN
        Next r
    End If
End Sub

'===========================
' ANOMALY HIGHLIGHTING
' - UNCONFIRMED cells → light yellow
' - EXCLUDED cells where ≥60% of other bidders are INCLUDED or have a dollar amount → orange
' - Skips exception/exclusion rows (rows starting with "Exception:")
'===========================
Private Sub HighlightAnomalies(ByVal ws As Worksheet)
    Dim anchor As Long: anchor = FindRowKeyInColB(ws, KEY_ADD_SCOPE)
    If anchor = 0 Then Exit Sub

    Dim lastCol As Long
    lastCol = ws.Cells(ROW_WIZ_ACTION, ws.Columns.Count).End(xlToLeft).Column

    ' Collect active bidder columns into an array
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

        ' Skip exception/exclusion rows (already highlighted red)
        If Left$(UCase$(desc), 10) = "EXCEPTION:" Then GoTo NextRow

        ' Count included vs excluded responses across all bidders
        Dim incCount As Long: incCount = 0
        For i = 1 To nBid
            Dim cv As String: cv = UCase$(Trim$(CStr(ws.Cells(r, bidCols(i)).Value)))
            If cv = TXT_INCLUDED Or IsNumericResponse(cv) Then incCount = incCount + 1
        Next i

        ' Apply per-cell highlighting
        For i = 1 To nBid
            Dim cellVal As String: cellVal = UCase$(Trim$(CStr(ws.Cells(r, bidCols(i)).Value)))

            If cellVal = TXT_UNCONF Then
                ws.Cells(r, bidCols(i)).Interior.Color = RGB_LIGHT_YELLOW

            ElseIf cellVal = TXT_EXCLUDED Then
                ' Orange if this bidder excluded something the majority included
                ' Only meaningful when there are 2+ bidders
                If nBid >= 2 And incCount / nBid >= 0.6 Then
                    ws.Cells(r, bidCols(i)).Interior.Color = RGB_ORANGE
                Else
                    ' Remove stale orange if situation has changed
                    If ws.Cells(r, bidCols(i)).Interior.Color = RGB_ORANGE Then
                        ws.Cells(r, bidCols(i)).Interior.ColorIndex = xlNone
                    End If
                End If

            Else
                ' Remove stale yellow or orange from cells that now have real data
                Dim curColor As Long: curColor = ws.Cells(r, bidCols(i)).Interior.Color
                If curColor = RGB_LIGHT_YELLOW Or curColor = RGB_ORANGE Then
                    ws.Cells(r, bidCols(i)).Interior.ColorIndex = xlNone
                End If
            End If
        Next i

NextRow:
    Next r
End Sub

' Returns True if the cell value looks like a dollar amount (numeric or formula result)
Private Function IsNumericResponse(ByVal s As String) As Boolean
    IsNumericResponse = IsNumeric(s) And Len(s) > 0
End Function
